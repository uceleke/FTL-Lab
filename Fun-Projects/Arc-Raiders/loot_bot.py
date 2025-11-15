import os
import discord
from discord import app_commands
from discord.ext import commands
import requests

# -------- Env vars --------
DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")
N8N_WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL")

print("[boot] DISCORD_TOKEN set:", bool(DISCORD_TOKEN))
print("[boot] N8N_WEBHOOK_URL:", N8N_WEBHOOK_URL)

if not DISCORD_TOKEN:
    raise RuntimeError("DISCORD_TOKEN is not set in environment.")
if not N8N_WEBHOOK_URL:
    raise RuntimeError("N8N_WEBHOOK_URL is not set in environment.")

# -------- Loot items for autocomplete (names only) --------
LOOT_ITEMS = [
    # Quests
    "Leaper Pulse Unit",
    "Power Rod",
    "Rocketeer Driver",
    "Surveyor Vault",
    "Antiseptic",
    "Hornet Driver",
    "Syringe",
    "Wasp Driver",
    "Water Pump",
    "Snitch Scanner",

    # Projects
    "Leaper Pulse Unit",
    "Magnetic Accelerator",
    "Exodus Modules",
    "Adv. Electrical Components",
    "Humidifier",
    "Sensors",
    "Cooling Fan",
    "Wires",
    "Durable Cloth",
    "Steel Spring",
    "Scrap Alloy",
    "Rubber Parts",
    "Metal Parts",
    "Battery",
    "Light Bulb",
    "Electrical Components",

    # Recycle
    "Accordion",
    "Alarm Clock",
    "ARC Coolant",
    "ARC Flex Rubber",
    "ARC Performance Steel",
    "ARC Synthetic Resin",
    "ARC Thermo Lining",
    "Bicycle Pump",
    "Broken Flashlight",
    "Broken Guidance System",
    "Broken Handcuffs",
    "Broken Handheld Radio",
    "Broken Taser",
    "Burned-out ARC Circuitry",
    "Camera Lens",
    "Candle Holder",
    "Coolant",
    "Cooling Coil",
    "Crumpled Plastic Bottle",
    "Damaged ARC Motion Core",
    "Damaged ARC Powercell",
    "Deflated Football",
    "Diving Goggles",
    "Dried-out ARC Resin",
    "Expired Respirator",
    "Flute",
    "Frying Pan",
    "Garlic Press",
    "Headphones",
    "Ice Cream Scooper",
    "Household Cleaner",
    "Impure ARC Coolant",
    "Industrial Charger",
    "Industrial Magnet",
    "Metal Brackets",
    "Number Plate",
    "Polluted Air Filter",
    "Radio",
    "Remote Control",
    "Ripped Safety Vest",
    "Rubber Pad",
    "Ruined Accordion",
    "Ruined Baton",
    "Ruined Handcuffs",
    "Ruined Parachute",
    "Ruined Riot Shield",
    "Ruined Tactical Vest",
    "Rusted Bolts",
    "Rusty ARC Steel",
    "Spotter Relay",
    "Spring Cushion",
    "Tattered ARC Lining",
    "Tattered Clothes",
    "Thermostat",
    "Torn Blanket",
    "Turbo Pump",
    "Water Filter",
]

# -------- Discord bot setup --------
intents = discord.Intents.default()  # no message_content needed for slash commands
bot = commands.Bot(command_prefix=None, intents=intents)


@bot.event
async def on_ready():
    """Sync slash commands globally so they work in any server."""
    try:
        await bot.tree.sync()
        print("[boot] Slash commands synced globally.")
    except Exception as e:
        print("[boot] Global slash sync error:", repr(e))
    print(f"[boot] Logged in as {bot.user} (ID: {bot.user.id})")


# -------- Autocomplete for /loot item --------
async def loot_autocomplete(
    interaction: discord.Interaction,
    current: str,
):
    """Return up to 25 loot name suggestions matching what the user typed."""
    current_lower = current.lower()

    # simple partial, case-insensitive match
    matches = [
        name for name in LOOT_ITEMS
        if current_lower in name.lower()
    ]

    # Discord allows max 25 choices
    matches = matches[:25]

    return [
        app_commands.Choice(name=name, value=name)
        for name in matches
    ]


# -------- /loot command --------
@bot.tree.command(name="loot", description="Check Arc Raiders loot info")
@app_commands.describe(item="Name of the item to look up")
@app_commands.autocomplete(item=loot_autocomplete)
async def loot(interaction: discord.Interaction, item: str):
    """
    Slash command:
    /loot <item>
    Sends loot lookup request to n8n and replies with formatted text from n8n.
    """

    guild_id = str(interaction.guild_id) if interaction.guild_id else None
    guild_name = interaction.guild.name if interaction.guild else None

    payload = {
        "item": item,
        "content": f"/loot {item}",
        "channel_id": str(interaction.channel_id),
        "author_id": str(interaction.user.id),
        "username": interaction.user.name,
        "guild_id": guild_id,
        "guild_name": guild_name,
    }

    try:
        r = requests.post(N8N_WEBHOOK_URL, json=payload, timeout=10)
        status = r.status_code
        print(f"[loot] n8n HTTP {status}")
        print(f"[loot] n8n response (first 200 chars): {r.text[:200]}")

        r.raise_for_status()
        data = r.json()

        # n8n usually returns a list of items; handle both list/dict
        if isinstance(data, list) and data:
            item_json = data[0].get("json", data[0])
        elif isinstance(data, dict):
            item_json = data.get("json", data)
        else:
            await interaction.response.send_message(
                "⚠️ Unexpected loot response format from n8n."
            )
            return

        result = item_json.get(
            "reply",
            "⚠️ Loot system did not include a reply field.",
        )

    except requests.exceptions.RequestException as e:
        print("[loot] Request error:", repr(e))
        result = "⚠️ Error reaching loot system (network error)."

    except Exception as e:
        print("[loot] Unexpected error:", repr(e))
        result = "⚠️ Error processing loot response."

    await interaction.response.send_message(result)


# -------- Run the bot --------
bot.run(DISCORD_TOKEN)