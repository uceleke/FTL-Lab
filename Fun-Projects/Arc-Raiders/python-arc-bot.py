import os
import logging
import discord
import requests

# --- Config from environment variables ---
DISCORD_TOKEN = os.getenv("DISCORD_TOKEN")
N8N_WEBHOOK_URL = os.getenv("N8N_WEBHOOK_URL")

if not DISCORD_TOKEN or not N8N_WEBHOOK_URL:
    raise SystemExit("ERROR: DISCORD_TOKEN and N8N_WEBHOOK_URL environment variables must be set.")

# --- Logging setup ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("arc-loot-bot")

# --- Discord client setup ---
intents = discord.Intents.default()
intents.message_content = True  # required to read message content

client = discord.Client(intents=intents)


@client.event
async def on_ready():
    log.info(f"Logged in as {client.user} (ID: {client.user.id})")
    log.info("Bot is ready and listening for !loot-bot commands.")


@client.event
async def on_message(message: discord.Message):
    # Ignore any bot messages (including ourselves)
    if message.author.bot:
        return

    content = message.content.strip()

    # Only respond to !loot-bot commands
    if not content.lower().startswith("!loot-bot"):
        return

    # Build payload for n8n
    payload = {
        "content": content,
        "channel_id": str(message.channel.id),
        "author_id": str(message.author.id),
        "username": message.author.name,
    }

    log.info(f"Received loot request from {message.author} in #{message.channel}: {content}")

    # Call n8n webhook
    try:
        resp = requests.post(N8N_WEBHOOK_URL, json=payload, timeout=8)
        resp.raise_for_status()

        data = resp.json()

        # n8n Webhook with responseMode=lastNode usually returns a list of items
        if isinstance(data, list) and len(data) > 0 and isinstance(data[0], dict):
            item = data[0].get("json", data[0])
        elif isinstance(data, dict):
            item = data.get("json", data)
        else:
            log.warning(f"Unexpected n8n response shape: {data}")
            await message.channel.send("⚠️ Loot service responded with an unexpected format.")
            return

        reply_text = item.get("reply")
        if not reply_text:
            reply_text = "⚠️ Loot service did not provide a reply field."

    except requests.exceptions.RequestException as e:
        log.error(f"Error calling n8n webhook: {e}")
        reply_text = "⚠️ Error talking to loot service. Try again in a bit."

    # Send reply back to Discord channel
    try:
        await message.channel.send(reply_text)
    except Exception as e:
        log.error(f"Error sending message to Discord: {e}")


if __name__ == "__main__":
    client.run(DISCORD_TOKEN)