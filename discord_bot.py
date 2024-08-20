import os
import asyncio
from functools import partial
import json
import requests
import discord
from discord.ext import commands
from dotenv import load_dotenv
from time import sleep
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from pywhatkit import sendwhatmsg_to_group_instantly
import logging

# Initialize logging
logging.basicConfig(level=logging.INFO)

# Load environment variables from .env file
load_dotenv()

# Setup for Discord bot
DISCORD_TOKEN = os.getenv('DISCORD_BOT_TOKEN')
TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
TELEGRAM_GROUP_ID = os.getenv('TELEGRAM_GROUP_ID')
WHATSAPP_GROUP_ID = os.getenv('WHATSAPP_GROUP_ID')
DISCORD_CHANNEL_ID = os.getenv('DISCORD_CHANNEL_ID')
GOOGLE_SHEET_ID = os.getenv('GOOGLE_SHEET_ID')

# Google Sheets API setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDS_FILE = 'credentials.json'

def authenticate_google_sheets():
    creds = None
    token_file = 'token.json'

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_file, 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    return service.spreadsheets().values()

sheet_service = authenticate_google_sheets()

# Setup for Discord bot
intents = discord.Intents.default()
intents.message_content = True

bot = commands.Bot(command_prefix='/', intents=intents)

current_poll_message = None
poll_options = []
poll_votes = {}
poll_active = False

@bot.event
async def on_ready():
    logging.info(f'Bot is ready. Logged in as {bot.user.name}')

@bot.command(name='create')
async def create_poll(ctx, *, question):
    global current_poll_message, poll_options, poll_votes, poll_active
    if poll_active:
        await ctx.send("A poll is already active. Please end it before creating a new one.")
        await ctx.message.delete()
        return

    poll_options = []
    poll_votes = {}
    poll_active = True
    embed = discord.Embed(title=question, description="React to vote!")
    current_poll_message = await ctx.send(embed=embed)
    await ctx.send("Use /add_option <option> to add poll options.")
    await ctx.message.delete()

@bot.command(name='add')
async def add_option(ctx, *, option):
    global poll_options, poll_active, current_poll_message
    if not poll_active:
        await ctx.send("No active poll to add options to.")
        await ctx.message.delete()
        return

    if len(poll_options) >= 20:
        await ctx.send("You can't add more than 20 options (Discord emoji limit).")
        await ctx.message.delete()
        return

    poll_options.append(option)
    emoji = f"{len(poll_options)}️⃣"
    await current_poll_message.add_reaction(emoji)
    embed = current_poll_message.embeds[0]
    embed.add_field(name=emoji, value=option, inline=False)
    await current_poll_message.edit(embed=embed)
    await ctx.message.delete()

@bot.event
async def on_reaction_add(reaction, user):
    global poll_votes, poll_active, current_poll_message
    if not poll_active or user.bot:
        return

    if reaction.message.id != current_poll_message.id:
        return

    option_index = int(reaction.emoji[0])
    poll_votes[user] = option_index
    update_google_sheet(user.name, poll_options[option_index - 1])

@bot.event
async def on_reaction_remove(reaction, user):
    global poll_votes, poll_active, current_poll_message
    if not poll_active or user.bot:
        return

    if reaction.message.id == current_poll_message.id:
        poll_votes.pop(user, None)
        update_google_sheet(user.name, None)

@bot.command(name='end')
async def end_poll(ctx):
    global poll_active
    if not poll_active:
        await ctx.send("No active poll to end.")
        await ctx.message.delete()
        return

    poll_active = False
    await ctx.send("Poll ended. Use /show_poll_result to see the results.")
    await ctx.message.delete()

@bot.command(name='result')
async def show_poll_result(ctx):
    if poll_active:
        await ctx.send("Poll is still active. End it before viewing the results.")
        await ctx.message.delete()
        return

    if not poll_options or not poll_votes:
        await ctx.send("No poll results to show.")
        await ctx.message.delete()
        return

    result_count = [0] * len(poll_options)
    for vote in poll_votes.values():
        result_count[vote - 1] += 1

    total_votes = len(poll_votes)
    if total_votes == 0:
        await ctx.send("No votes have been cast.")
        await ctx.message.delete()
        return

    result_percentages = [(i + 1, (count / total_votes) * 100) for i, count in enumerate(result_count)]
    sorted_options = sorted(result_percentages, key=lambda item: item[1], reverse=True)

    result_message = "**Poll Results:**\n\n"
    for option_number, percentage in sorted_options:
        result_message += f"{poll_options[option_number - 1]}: {percentage:.2f}%\n"

    result_message += "\n**Voters and their choices:**\n"
    for user, option_number in poll_votes.items():
        result_message += f"{user.name}: {poll_options[option_number - 1]}\n"

    # Print the results in Discord
    await ctx.send(result_message)

    # Send results to Telegram and WhatsApp
    if result_message:
        try:
            await send_alert_telegram(result_message)
        except Exception as e:
            logging.error(f"Failed to send Telegram alert: {str(e)}")
        # sleep(2)
        # try:
        #     await send_alert_whatsapp(result_message)
        # except Exception as e:
        #     logging.error(f"Failed to send WhatsApp alert: {str(e)}")

def update_google_sheet(username, preferred_item):
    try:
        # Read existing data
        sheet_data = sheet_service.get(spreadsheetId=GOOGLE_SHEET_ID, range="A:C").execute().get('values', [])

        # Find existing row for the user
        user_row = None
        for i, row in enumerate(sheet_data):
            if row[1] == username:
                user_row = i + 1  # Sheet row numbers start at 1

        # Prepare data to write
        if preferred_item:
            timestamp = discord.utils.utcnow().strftime('%Y-%m-%d %H:%M:%S')
            new_row = [timestamp, username, preferred_item]
            if user_row:
                # Update existing row
                sheet_service.update(spreadsheetId=GOOGLE_SHEET_ID, range=f"A{user_row}:C{user_row}", valueInputOption='RAW', body={'values': [new_row]}).execute()
            else:
                # Append new row
                sheet_service.append(spreadsheetId=GOOGLE_SHEET_ID, range="A:C", valueInputOption='RAW', insertDataOption='INSERT_ROWS', body={'values': [new_row]}).execute()
        else:
            # Clear the user's row if they removed their selection
            if user_row:
                sheet_service.update(spreadsheetId=GOOGLE_SHEET_ID, range=f"A{user_row}:C{user_row}", valueInputOption='RAW', body={'values': [["", "", ""]]}).execute()

    except Exception as e:
        logging.error(f"Failed to update Google Sheet: {str(e)}")

def send_alert_telegram(message_text):
    try:
        telegram_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
        telegram_params = {'chat_id': TELEGRAM_GROUP_ID, 'text': message_text}
        requests.post(telegram_url, data=telegram_params)
    except Exception as e:
        logging.error(f"Failed to send Telegram message: {str(e)}")

async def send_alert_whatsapp(message_text):
    try:
        loop = asyncio.get_running_loop()
        await loop.run_in_executor(None, partial(sendwhatmsg_to_group_instantly, WHATSAPP_GROUP_ID, message_text))
    except Exception as e:
        logging.error(f"Failed to send WhatsApp message: {str(e)}")

# Run the bot
bot.run(DISCORD_TOKEN)
