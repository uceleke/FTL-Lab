<# 
Password Expiring Notification Script 
- Created to be used with a GPO - scheduled task
- First notification will be seen at 14 days. Second at 7 and then everyday after that till password expires. 
- Can deploy using Network share that has at least read access by all users or in sysvol folder. 
OK-only dialog for password expiry with logging.
- No RSAT/modules required
- Self-relaunches in STA under Windows PowerShell if needed
- Logs to file (per-user) and best-effort to Windows Event Log

Created by - Jesse Esposo 20251104 - FTL - Big Dawg
#>

[CmdletBinding()]
param(
    [switch]$ForceTest,
    [int]$SimulatedDaysLeft,
    [switch]$LogToFile,                  # enable file logging explicitly (recommended in GPO)
    [string]$LogPath,                    # custom log path (optional). Default shown below
    [string]$OrgName = 'Contoso',        # appears in window title and event source
    [int]$LogRetentionDays = 14          # old logs auto-rotated
)

# ================= SETTINGS =================
$ForceOnOrAfterExpiry = $true           # also warn when daysLeft <= 0
$DefaultLogPath = Join-Path $env:LOCALAPPDATA "PwdExpiry\PwdExpiryDialog.log"

# ================= LOGGING HELPERS =================
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')] [string]$Level = 'INFO'
    )
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
    $line = "[$timestamp] [$Level] $Message"

    # File logging
    if ($script:FileLoggingEnabled) {
        try {
            $dir = Split-Path -Parent $script:LogFile
            if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Add-Content -Path $script:LogFile -Value $line
        } catch { }
    }

    # Event log (best-effort, non-fatal)
    if ($Level -eq 'ERROR' -or $Level -eq 'WARN') {
        try {
            $src = "$OrgName-PwdExpiry"
            if (-not [System.Diagnostics.EventLog]::SourceExists($src)) {
                # Requires admin; wrap in try so failure doesn't break script
                New-EventLog -LogName Application -Source $src -ErrorAction Stop
            }
            $etype = if ($Level -eq 'ERROR') { 'Error' } else { 'Warning' }
            Write-EventLog -LogName Application -Source $src -EventId 31001 -EntryType $etype -Message $Message
        } catch { }
    }
}

function Invoke-LogRotation {
    try {
        if (Test-Path $script:LogFile) {
            $fi = Get-Item $script:LogFile
            if ($fi.Length -gt 2MB) {
                $archive = Join-Path (Split-Path -Parent $script:LogFile) ("PwdExpiryDialog_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                Move-Item -Path $script:LogFile -Destination $archive -Force
            }
            # prune old logs
            Get-ChildItem (Split-Path -Parent $script:LogFile) -Filter 'PwdExpiryDialog_*.log' -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-$LogRetentionDays) } |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

# Initialize logging
$script:LogFile = if ($PSBoundParameters.ContainsKey('LogPath') -and $LogPath) { $LogPath } else { $DefaultLogPath }
$script:FileLoggingEnabled = $LogToFile.IsPresent -or $PSBoundParameters.ContainsKey('LogPath')

if ($script:FileLoggingEnabled) { Invoke-LogRotation }

Write-Log -Message "=== Script start: user=$env:USERDOMAIN\$env:USERNAME, pid=$PID, host=$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion), forceTest=$ForceTest, simDays=$SimulatedDaysLeft ===" -Level 'DEBUG'

# ================= RUNTIME/STA SAFETY =================
function Test-IsSTA { [Threading.Thread]::CurrentThread.GetApartmentState() -eq 'STA' }
function Invoke-InSTA {
    param([string]$ArgsLine)
    $psExe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path $psExe)) { $psExe = "powershell.exe" }
    $scriptPath = $MyInvocation.MyCommand.Path
    $staArgs = @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-File',('"{0}"' -f $scriptPath))
    if ($ArgsLine) { $staArgs += $ArgsLine }
    Write-Log -Message "Relaunching self in STA via $psExe $ArgsLine" -Level 'DEBUG'
    Start-Process -FilePath $psExe -ArgumentList $staArgs -WindowStyle Normal | Out-Null
    exit
}

# Relaunch in STA on Windows PowerShell if needed (WPF reliability)
$hostIsPwsh = ($PSVersionTable.PSEdition -eq 'Core')
if (-not (Test-IsSTA) -or $hostIsPwsh) {
    $passArgs = @()
    if ($ForceTest) { $passArgs += '-ForceTest' }
    if ($script:FileLoggingEnabled) { $passArgs += '-LogToFile' }
    if ($PSBoundParameters.ContainsKey('LogPath')) { $passArgs += ('-LogPath "{0}"' -f $script:LogFile) }
    if ($PSBoundParameters.ContainsKey('SimulatedDaysLeft')) { $passArgs += "-SimulatedDaysLeft $SimulatedDaysLeft" }
    $passArgs += ('-OrgName "{0}"' -f $OrgName)
    Invoke-InSTA -ArgsLine ($passArgs -join ' ')
}

# ================= DOMAIN/EXPIRY =================
function Convert-LargeIntegerToInt64 {
    param([Parameter(Mandatory)]$LargeInteger)
    try {
        $high = [int64]$LargeInteger.HighPart
        $low  = [uint32]$LargeInteger.LowPart
        return ($high -shl 32) -bor $low
    } catch { return $null }
}

function Get-MaxPwdAgeFromADSI {
    try {
        $root = [ADSI]'LDAP://RootDSE'
        $namingCtx = $root.defaultNamingContext
        if (-not $namingCtx) { return $null }
        $domainDE = [ADSI]("LDAP://{0}" -f $namingCtx)
        $li = $domainDE.Get("maxPwdAge")
        if (-not $li) { return $null }
        $ticks = Convert-LargeIntegerToInt64 $li
        if ($null -eq $ticks) { return $null }
        if ($ticks -eq 0) { return [TimeSpan]::Zero } # never expires
        return [TimeSpan]::FromTicks($ticks) # typically negative
    } catch { return $null }
}

function Get-MaxPwdAgeFromNetAccounts {
    try {
        $out = (cmd /c 'net accounts /domain') 2>$null
        if (-not $out) { return $null }
        $line = $out | Where-Object { $_ -match '(?i)Maximum\s+password\s+age' }
        if (-not $line) { return $null }
        if ($line -match '(?i)(Unlimited|Not\s*Set|Never)') { return [TimeSpan]::Zero }
        if ($line -match '([0-9]+)\s*day') {
            return [TimeSpan]::FromDays([int]$matches[1])
        }
        if ($line -match '([0-9]+)\s*$') {
            return [TimeSpan]::FromDays([int]$matches[1])
        }
        return $null
    } catch { return $null }
}

function ConvertTo-DateLoose {
    param([Parameter(Mandatory)][string]$Text)
    try {
        return [datetime]::Parse($Text, [System.Globalization.CultureInfo]::CurrentCulture)
    } catch {
        try {
            return [datetime]::Parse($Text, [System.Globalization.CultureInfo]::InvariantCulture)
        } catch { return $null }
    }
}

function Get-UserPwdExpiryDays {
    [CmdletBinding()]
    param([switch]$VerboseLog)

    # 0) quick off-ramps
    if (-not $env:USERDNSDOMAIN -or [string]::IsNullOrWhiteSpace($env:USERDNSDOMAIN)) {
        if ($VerboseLog) { Write-Host "[INFO] Not domain-joined (USERDNSDOMAIN empty)"; }
        return $null
    }

    # 1) First try: net user /domain (reads “Password expires” directly)
    try {
        $nu = (cmd /c ("net user {0} /domain" -f $env:USERNAME)) 2>$null
        if ($nu -and $nu.Count -gt 0) {
            if ($VerboseLog) { Write-Host "[INFO] net user /domain output found." }
            # Lines often look like:
            # "Password last set             10/31/2025 3:14 PM"
            # "Password expires              11/04/2025 3:14 PM"
            $lastSetLine  = $nu | Where-Object { $_ -match '^(?i)\s*Password\s+last\s+set' } | Select-Object -First 1
            $expiresLine  = $nu | Where-Object { $_ -match '^(?i)\s*Password\s+expires' }   | Select-Object -First 1
            $neverFlag    = $expiresLine -and ($expiresLine -match '(?i)Never|Unlimited|Not\s*Set')

            if ($expiresLine) {
                if ($neverFlag) {
                    if ($VerboseLog) { Write-Host "[INFO] Password expires: NEVER (user or policy)"; }
                    return $null
                }
                # take substring after the label
                $expText = ($expiresLine -replace '^(?i)\s*Password\s+expires\s+','').Trim()
                $expires = ConvertTo-DateLoose-DateLoose $expText
                if ($expires) {
                    $days = [math]::Floor(($expires - (Get-Date)).TotalDays)
                    if ($VerboseLog) { Write-Host "[OK] Expiry from net user: $expires ($days days)"; }
                    return [pscustomobject]@{ DisplayName=$env:USERNAME; ExpiresOn=$expires; DaysLeft=$days; Source='net user' }
                }
            }
        }
    } catch { if ($VerboseLog) { Write-Host "[WARN] net user parse failed: $($_.Exception.Message)"; } }

    # 2) Second try: ADSI maxPwdAge + UserPrincipal.LastPasswordSet
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
        $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current
        if ($user -and $user.PasswordNeverExpires) {
            if ($VerboseLog) { Write-Host "[INFO] User flag PasswordNeverExpires=True"; }
            return $null
        }

        $maxAge = Get-MaxPwdAgeFromADSI
        if ($maxAge) {
            if ($maxAge.Ticks -eq 0) {
                if ($VerboseLog) { Write-Host "[INFO] Domain maxPwdAge=0 (never expires)."; }
                return $null
            }
            $lastSet = if ($user.LastPasswordSet) { $user.LastPasswordSet } else { Get-Date }
            $expires = if ($maxAge.Ticks -lt 0) { $lastSet - $maxAge } else { $lastSet + $maxAge }
            $days    = [math]::Floor(($expires - (Get-Date)).TotalDays)
            if ($VerboseLog) { Write-Host "[OK] Expiry from ADSI: $expires ($days days)"; }
            return [pscustomobject]@{ DisplayName=$env:USERNAME; ExpiresOn=$expires; DaysLeft=$days; Source='ADSI' }
        }
    } catch { if ($VerboseLog) { Write-Host "[WARN] ADSI path failed: $($_.Exception.Message)"; } }

    # 3) Fallback: net accounts (days) + “Password last set” from net user
    try {
        $maxAge2 = Get-MaxPwdAgeFromNetAccounts
        if ($maxAge2) {
            if ($maxAge2.Ticks -eq 0) { if ($VerboseLog) { Write-Host "[INFO] net accounts shows never expires."; }; return $null }

            # Get "Password last set" from net user if available
            $lastSet = $null
            if ($nu -and $nu.Count -gt 0) {
                $lastSetLine = $nu | Where-Object { $_ -match '^(?i)\s*Password\s+last\s+set' } | Select-Object -First 1
                if ($lastSetLine) {
                    $lsText = ($lastSetLine -replace '^(?i)\s*Password\s+last\s+set\s+','').Trim()
                    $lastSet = ConvertTo-DateLoose $lsText
                }
            }
            if (-not $lastSet) {
                # final fallback: use UserPrincipal if we have it
                try {
                    if (-not $user) {
                        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction Stop
                        $user = [System.DirectoryServices.AccountManagement.UserPrincipal]::Current
                    }
                    $lastSet = if ($user -and $user.LastPasswordSet) { $user.LastPasswordSet } else { Get-Date }
                } catch { $lastSet = Get-Date }
            }

            $expires = $lastSet + $maxAge2
            $days    = [math]::Floor(($expires - (Get-Date)).TotalDays)
            if ($VerboseLog) { Write-Host "[OK] Expiry from net accounts + last set: $expires ($days days)"; }
            return [pscustomobject]@{ DisplayName=$env:USERNAME; ExpiresOn=$expires; DaysLeft=$days; Source='net accounts' }
        }
    } catch { if ($VerboseLog) { Write-Host "[WARN] net accounts fallback failed: $($_.Exception.Message)"; } }

    if ($VerboseLog) { Write-Host "[INFO] All methods failed. Likely offline to DC, locale mismatch, or policy says never expires." }
    return $null
}

# ================= UI (STRICT OK) =================
function Show-StrictOkDialog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [string]$Title = 'Password Expiring'
    )
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction Stop

        [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="$Title" Height="220" Width="520"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        WindowStyle="None"
        Topmost="True"
        ShowInTaskbar="True"
        Background="#FFFDFDFD">
  <Border Margin="12" Padding="16" CornerRadius="8" Background="White" BorderBrush="#FFCCCCCC" BorderThickness="1">
    <DockPanel LastChildFill="True">
      <TextBlock DockPanel.Dock="Top" TextWrapping="Wrap" FontSize="14" Margin="0,0,0,12">$Message</TextBlock>
      <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal" HorizontalAlignment="Right">
        <Button Name="OkBtn" Width="88" Height="28" Margin="8,0,0,0">OK</Button>
      </StackPanel>
    </DockPanel>
  </Border>
</Window>
"@

        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $window = [Windows.Markup.XamlReader]::Load($reader)

        $script:closedByOk = $false
        $null = $window.Add_Closing({
            param($wSender, $wArgs)
            if (-not $script:closedByOk) { $wArgs.Cancel = $true }  # block X/Alt+F4
        })

        $okBtn = $window.FindName('OkBtn')
        $okBtn.Add_Click({ $script:closedByOk = $true; $window.Close() })

        # Ensure a dispatcher exists (discard result to satisfy analyzers)
        if (-not [System.Windows.Threading.Dispatcher]::FromThread([System.Threading.Thread]::CurrentThread)) {
            [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher
        }

        [void]$window.ShowDialog()
        Write-Log -Message "UI shown via WPF strict dialog." -Level 'INFO'
        return $true
    } catch {
        Write-Log -Message ("WPF failed, falling back to WinForms: {0}" -f $_.Exception.Message) -Level 'WARN'
        try {
            Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
            [void][System.Windows.Forms.MessageBox]::Show(
                $Message, $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            Write-Log -Message "UI shown via WinForms MessageBox." -Level 'INFO'
            return $true
        } catch {
            Write-Log -Message ("WinForms also failed: {0}" -f $_.Exception.Message) -Level 'ERROR'
            return $false
        }
    }
}

# ================= MAIN =================
try {
    $info = Get-UserPwdExpiryDays
    if (-not $info) {
        if (-not $ForceTest -and -not $PSBoundParameters.ContainsKey('SimulatedDaysLeft')) {
            Write-Log -Message "No expiry info and not in ForceTest/simulate mode. Exiting quietly." -Level 'INFO'
            return
        }
        # synthesize a fake expiry for ForceTest
        $fakeExpires = (Get-Date).AddDays(5)
        $info = [PSCustomObject]@{ DisplayName=$env:USERNAME; ExpiresOn=$fakeExpires; DaysLeft=5 }
        Write-Log -Message "ForceTest path: synthesized expiry in 5 days." -Level 'DEBUG'
    }

    $days = if ($PSBoundParameters.ContainsKey('SimulatedDaysLeft')) { 
        Write-Log -Message "Using SimulatedDaysLeft=$SimulatedDaysLeft" -Level 'DEBUG'
        $SimulatedDaysLeft 
    } else { 
        $info.DaysLeft 
    }

    $shouldNotify =
        $ForceTest -or
        ($days -eq 14) -or
        ($days -eq 7)  -or
        ($days -ge 1 -and $days -le 6) -or
        ($ForceOnOrAfterExpiry -and $days -lt 1)

    Write-Log -Message ("Decision: daysLeft={0}, shouldNotify={1}, force={2}" -f $days, $shouldNotify, $ForceTest) -Level 'INFO'

    if (-not $shouldNotify) {
        Write-Log -Message "Outside of notification window. Exiting quietly." -Level 'INFO'
        return
    }

    $msg =
        if ($days -gt 1) {
            "Your password expires in $days days (on $($info.ExpiresOn.ToString('ddd, MMM d, yyyy h:mm tt'))). Please change it now: Ctrl+Alt+Del > Change a password."
        } elseif ($days -eq 1) {
            "Your password expires in 1 day (on $($info.ExpiresOn.ToString('ddd, MMM d, yyyy h:mm tt'))). Please change it now: Ctrl+Alt+Del > Change a password."
        } else {
            "Your password expires today ($($info.ExpiresOn.ToString('ddd, MMM d, yyyy h:mm tt'))). Please change it now: Ctrl+Alt+Del > Change a password."
        }

    [void](Show-StrictOkDialog -Message $msg -Title "$OrgName Password Expiring")
}
catch {
    Write-Log -Message ("Unhandled error: {0}`n{1}" -f $_.Exception.Message, $_.ScriptStackTrace) -Level 'ERROR'
}
finally {
    Write-Log -Message "=== Script end ===" -Level 'DEBUG'
}