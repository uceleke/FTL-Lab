<#
.SYNOPSIS
  Manually deploy Windows Updates from WSUS at scale with NO immediate reboot.
  Later, trigger a coordinated reboot only on machines that need it.

.PARAMETERS
  -Targets         : AD group, OU, or a text file list of machines.
  -ModuleSource    : UNC path to PSWindowsUpdate module folder (if not in PSGallery).
  -Throttle        : Parallelism for remoting.
  -StageOnly       : If set, only download (no install) to pre-stage content.
  -InstallNow      : Default. Download+Install updates, but DO NOT reboot.
  -ScheduleReboot  : Schedule a reboot at OffHoursTime **only for pending-reboot machines**.
  -OffHoursTime    : Local time on target (e.g. "02:30"); default 02:00.
  -RebootDelayMin  : Minutes of warning before scheduled reboot when user is logged on.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
  [Parameter(Mandatory=$true)]
  [string]$Targets,            # e.g. "OU=Workstations,DC=corp,DC=local" or "CN=WSUS-Pilot,OU=Groups,DC=corp,DC=local" or "C:\temp\hosts.txt"

  [Parameter(Mandatory=$true)]
  [string]$ModuleSource,       # e.g. "\\fileserver\PSPackages\PSWindowsUpdate"

  [int]$Throttle = 20,

  [switch]$StageOnly,
  [switch]$InstallNow,
  [switch]$ScheduleReboot,

  [string]$OffHoursTime = "02:00",
  [int]$RebootDelayMin = 15,

  [string]$OutputPath = ".\WSUS_ManualWave_$(Get-Date -Format yyyyMMdd_HHmm).csv"
)

function Resolve-Targets {
  param([string]$Specifier)
  if (Test-Path $Specifier) {
    Get-Content -Path $Specifier | Where-Object { $_ -and $_ -notmatch "^\s*#" } | Sort-Object -Unique
  }
  elseif ($Specifier -like "OU=*") {
    try {
      Import-Module ActiveDirectory -ErrorAction Stop
      Get-ADComputer -SearchBase $Specifier -Filter * | Select-Object -Expand Name | Sort-Object -Unique
    } catch {
      throw "Could not query AD OU. Ensure RSAT AD module is available. $_"
    }
  }
  else {
    # Treat as AD group
    try {
      Import-Module ActiveDirectory -ErrorAction Stop
      Get-ADGroupMember $Specifier -Recursive | Where-Object objectClass -eq "computer" | Select-Object -Expand Name | Sort-Object -Unique
    } catch {
      throw "Could not query AD group. Ensure RSAT AD module is available. $_"
    }
  }
}

$computers = Resolve-Targets -Specifier $Targets
if (-not $computers) { throw "No targets resolved from '$Targets'." }

Write-Host "Targets: $($computers.Count) computers."

# Scriptblock runs on each target
$perNode = {
  param(
    [string]$ModuleSource,
    [bool]$StageOnly,
    [bool]$InstallNow,
    [string]$OffHoursTime,
    [int]$RebootDelayMin
  )
  $result = [ordered]@{
    ComputerName     = $env:COMPUTERNAME
    Phase            = $null
    InstalledCount   = 0
    DownloadedCount  = 0
    PendingReboot    = $false
    HadErrors        = $false
    ErrorMessage     = $null
  }

  try {
    # Ensure PSWindowsUpdate exists
    $modulePath = Join-Path $env:ProgramFiles "WindowsPowerShell\Modules\PSWindowsUpdate"
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
      if (-not (Test-Path $ModuleSource)) { throw "ModuleSource not accessible: $ModuleSource" }
      New-Item -ItemType Directory -Force -Path $modulePath | Out-Null
      Copy-Item -Path (Join-Path $ModuleSource "*") -Destination $modulePath -Recurse -Force
    }
    Import-Module PSWindowsUpdate -ErrorAction Stop

    # Make sure Windows Update service is running
    $wus = Get-Service wuauserv -ErrorAction SilentlyContinue
    if ($wus.Status -ne 'Running') { Start-Service wuauserv }

    # Optional: trust WSUS policy already configured via GPO. We just act on it.

    if ($StageOnly) {
      $result.Phase = "StageOnly"
      # Download only
      $dl = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Download -IgnoreReboot -ErrorAction Stop
      $result.DownloadedCount = ($dl | Measure-Object).Count
    }
    elseif ($InstallNow) {
      $result.Phase = "InstallNoReboot"
      $inst = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -ErrorAction Stop
      $result.InstalledCount = ($inst | Measure-Object).Count
    }

    # Check pending reboot
    $reboot = Get-WURebootStatus
    $result.PendingReboot = [bool]$reboot.IsRebootRequired

    # Optionally schedule a controlled reboot if pending
    if ($InstallNow -and $using:ScheduleReboot -and $result.PendingReboot) {
      # Create a one-time scheduled task at OffHoursTime local
      $taskName = "ControlledReboot-WSUS"
      $time = [DateTime]::Today.Add([TimeSpan]::Parse($OffHoursTime))
      if ($time -lt (Get-Date)) { $time = $time.AddDays(1) } # tomorrow if time already passed

      $action  = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -WindowStyle Hidden -Command `"Start-Sleep -Seconds $($RebootDelayMin*60); Restart-Computer -Force`""
      $trigger = New-ScheduledTaskTrigger -Once -At $time
      $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
      try {
        # Clean previous if exists
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
          Unregister-ScheduledTask -TaskName $taskName -Confirm:$false | Out-Null
        }
      } catch {}
      Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal | Out-Null

      # Gentle heads-up to signed-in users (fallback if BurntToast isn't present)
      try {
        $sessions = (quser) -replace '\s{2,}', ',' | ConvertFrom-Csv -Header 'USER','STATE','ID','IDLE','LOGON'
        if ($sessions | Where-Object { $_.STATE -like 'Active*' }) {
          msg * /time:$([Math]::Max(10,[int]($RebootDelayMin*60))) "Updates installed. This device will reboot around $($time.ToShortTimeString()) unless you save work. IT scheduled restart to finish updates."
        }
      } catch {}

      $result.Phase = "$($result.Phase)+RebootScheduled@$($time.ToString('yyyy-MM-dd HH:mm'))"
    }

  } catch {
    $result.HadErrors = $true
    $result.ErrorMessage = $_.Exception.Message
  }

  # Minimal per-node return
  [pscustomobject]$result
}

$jobs = @()
$computers | ForEach-Object {
  $jobs += Invoke-Command -ComputerName $_ -ScriptBlock $perNode -ArgumentList $ModuleSource, $StageOnly.IsPresent, $InstallNow.IsPresent, $OffHoursTime, $RebootDelayMin -AsJob -ErrorAction SilentlyContinue
}

Write-Host "Dispatched $($jobs.Count) jobs…"
$all = Receive-Job -Job $jobs -Wait -AutoRemoveJob

# Save report
$all | Sort-Object ComputerName | Export-Csv -NoTypeInformation -Path $OutputPath
Write-Host "Report: $OutputPath"
$all