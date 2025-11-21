import discord
from discord.ext import commands, tasks
import logging
from dotenv import load_dotenv
import os
import requests
import asyncio
import aiohttp
from aiohttp import web
import threading
import json
import time
import csv
import requests
import os
import re
import logging
import zipfile

from pathlib import Path

from datetime import datetime

import json
import time
import csv
import requests
import os

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

from datetime import datetime

load_dotenv()
token = os.getenv('DISCORD_TOKEN')

# Configure logging for both local and Render environments
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),  # This outputs to stdout/stderr for Render logs
        logging.FileHandler('hopiumbot.log', encoding='utf-8', mode='a') if not os.getenv('RENDER') else logging.NullHandler()
    ]
)

# Store ongoing applications
active_applications = {}

# Set discord.py logging level to reduce spam
logging.getLogger('discord').setLevel(logging.WARNING)
logging.getLogger('discord.http').setLevel(logging.WARNING)

# Create logger for our bot
logger = logging.getLogger('HopiumBot')
intents = discord.Intents.default()
intents.message_content = True  # Enable message content intent
intents.guilds = True  # Enable guild intents
intents.members = True  # Enable member intents

# Application questions
APPLICATION_QUESTIONS = [
    "Character name:",
    "Class/Spec:",
    "What country are you from and how old are you?",
    "Please tell us a bit about yourself, who are you outside of the game?",
    "Explain your WoW experience. Include logs of past relevant characters (Classic/SoM//SoD/Retail).",
    "We require a few things from every raider in the guild. To have above average performance for your class and atleast 80% raid attendance. Can you fulfill these requirements?",
    "With TBC just around the corner, which class/spec are you planning to play? Are you planning to play any alts?",
    "Why did you choose to apply to <Hopium>?",
    "Can someone in <Hopium> vouch for you?",
    "Surprise us! What's something you'd like to tell us, it can be absolutely anything!"
]

CLASS_LIST = {
    "Druid" : {"name": "Druid", "roles": ["DPS", "Heal", "Tank"], "color": "FF7C0A"},
    "Hunter" : {"name": "Hunter", "roles": ["DPS"], "color": "AAD372"},
    "Mage" : {"name": "Mage", "roles": ["DPS"], "color": "3FC7EB"},
    "Paladin" : {"name": "Paladin", "roles": ["DPS", "Heal", "Tank"], "color": "F48CBA"},
    "Priest" : {"name": "Priest", "roles": ["Heal", "DPS"], "color": "FFFFFF"},
    "Rogue" : {"name": "Rogue", "roles": ["DPS"], "color": "FFF468"},
    "Shaman" : {"name": "Shaman", "roles": ["DPS", "Heal"], "color": "0070DD"},
    "Warlock" : {"name": "Warlock", "roles": ["DPS"], "color": "8788EE"},
    "Warrior" : {"name": "Warrior", "roles": ["DPS", "Tank"], "color": "C69B6D"}
}

BLIZZARD_ID = os.getenv('BLIZZARD_ID')
BLIZZARD_SECRET = os.getenv('BLIZZARD_SECRET')
BLIZZARD_TOKEN_URL = 'https://eu.battle.net/oauth/token'

WCL_ID = os.getenv('WCL_ID')
WCL_SECRET = os.getenv('WCL_SECRET')

# Validate required environment variables
missing_vars = []
if not BLIZZARD_ID:
    missing_vars.append('BLIZZARD_ID')
if not BLIZZARD_SECRET:
    missing_vars.append('BLIZZARD_SECRET')
if not WCL_ID:
    missing_vars.append('WCL_ID')
if not WCL_SECRET:
    missing_vars.append('WCL_SECRET')

if missing_vars:
    error_msg = f"Missing required environment variables: {', '.join(missing_vars)}"
    print(f"❌ {error_msg}")
    logger.error(error_msg)
    print("Please set these variables in your .env file or Render environment variables.")
else:
    logger.info("All required API credentials loaded successfully")

bot = commands.Bot(command_prefix='!', intents=intents)

role = "Trial"

# Define guild-specific paths that work both locally and on Render
def get_guild_data_path(guild_id):
    """Get data path specific to a guild"""
    if os.getenv('RENDER'):
        # Production on Render
        base_path = '/app/data'
    else:
        # Local development
        base_dir = os.path.dirname(os.path.abspath(__file__))
        base_path = os.path.join(base_dir, 'app', 'data')
    
    # Create guild-specific directory
    guild_path = os.path.join(base_path, f'guild_{guild_id}')
    os.makedirs(guild_path, exist_ok=True)
    return guild_path

def get_guild_file_paths(guild_id):
    """Get all file paths for a specific guild"""
    guild_data_dir = get_guild_data_path(guild_id)
    
    # TMB directory and files
    tmb_dir = os.path.join(guild_data_dir, 'tmb')
    character_file = os.path.join(tmb_dir, 'character-json.json')
    attendance_file = os.path.join(tmb_dir, 'hopium-attendance.csv')
    item_file = os.path.join(tmb_dir, 'item-notes.csv')
    
    # Cache directory and files
    cache_dir = os.path.join(guild_data_dir, 'cache')
    armory_file = os.path.join(cache_dir, 'armory.json')
    item_icons_file = os.path.join(cache_dir, 'item-icons.json')
    parses_file = os.path.join(cache_dir, 'parses.json')
    
    # Sheet directory
    sheet_dir = os.path.join(guild_data_dir, 'sheets')
    
    return {
        'guild_data_dir': guild_data_dir,
        'tmb_dir': tmb_dir,
        'cache_dir': cache_dir,
        'sheet_dir': sheet_dir,
        'character_file': character_file,
        'attendance_file': attendance_file,
        'item_file': item_file,
        'armory_file': armory_file,
        'item_icons_file': item_icons_file,
        'parses_file': parses_file
    }

def calculate_gradient_color(value, start_color, end_color):
    value = max(0, min(1, value))

    start_red, start_green, start_blue = start_color
    end_red, end_green, end_blue = end_color

    red = int(start_red + (end_red - start_red) * value)
    green = int(start_green + (end_green - start_green) * value)
    blue = int(start_blue + (end_blue - start_blue) * value)

    return f"{red:02X}{green:02X}{blue:02X}"

# Initialize guild-specific data files
def initialize_guild_data_files(guild_id):
    """Initialize required directories and files for a specific guild"""
    try:
        paths = get_guild_file_paths(guild_id)
        
        # Create directories
        os.makedirs(paths['tmb_dir'], exist_ok=True)
        os.makedirs(paths['cache_dir'], exist_ok=True)
        os.makedirs(paths['sheet_dir'], exist_ok=True)
        
        # Initialize character file if it doesn't exist
        if not os.path.exists(paths['character_file']):
            with open(paths['character_file'], 'w', encoding='utf-8') as f:
                json.dump([], f)
            print(f"📄 Created empty character file for guild {guild_id}")
        
        # Initialize armory file if it doesn't exist
        if not os.path.exists(paths['armory_file']):
            with open(paths['armory_file'], 'w', encoding='utf-8') as f:
                json.dump({}, f)
            print(f"📄 Created empty armory file for guild {guild_id}")
        
        # Initialize item icons file if it doesn't exist
        if not os.path.exists(paths['item_icons_file']):
            with open(paths['item_icons_file'], 'w', encoding='utf-8') as f:
                json.dump({}, f)
            print(f"📄 Created empty item icons file for guild {guild_id}")
        
        # Initialize parses file if it doesn't exist
        if not os.path.exists(paths['parses_file']):
            with open(paths['parses_file'], 'w', encoding='utf-8') as f:
                json.dump({}, f)
            print(f"📄 Created empty parses file for guild {guild_id}")
        
        print(f"✅ Data files initialized successfully for guild {guild_id}")
        logger.info(f"Data files initialized successfully for guild {guild_id}")
        
    except Exception as e:
        print(f"❌ Error initializing data files for guild {guild_id}: {e}")
        logger.error(f"Error initializing data files for guild {guild_id}: {e}", exc_info=True)

# Background task that runs every X minutes
@tasks.loop(minutes=15)  # Reduced frequency - 15 minutes is more suitable for API calls
async def periodic_task():
    try:
        current_time = datetime.now().strftime('%H:%M:%S')
        print(f"🔄 Starting periodic update for all guilds at {current_time}")
        logger.info(f"Starting periodic update for all guilds at {current_time}")
        
        # Process each guild the bot is in
        for guild in bot.guilds:
            try:
                print(f"🏰 Processing guild: {guild.name} (ID: {guild.id})")
                logger.info(f"Processing guild: {guild.name} (ID: {guild.id})")
                
                # Initialize guild data files if needed
                initialize_guild_data_files(guild.id)
                
                # Get guild-specific file paths
                paths = get_guild_file_paths(guild.id)
                
                # Check if required files exist for this guild
                if not os.path.exists(paths['character_file']):
                    print(f"⚠️ Character file not found for guild {guild.name}: {paths['character_file']}")
                    continue
                
                # Process this guild's data
                await process_guild_data(guild.id, paths)
                
            except Exception as e:
                print(f"❌ Error processing guild {guild.name}: {e}")
                logger.error(f"Error processing guild {guild.name}: {e}", exc_info=True)
                continue
        
        print(f"✅ Periodic update completed for all guilds")
        logger.info(f"Periodic update completed for all guilds")
        
    except Exception as e:
        print(f"❌ Critical error in periodic task: {e}")
        logger.error(f"Critical error in periodic task: {e}", exc_info=True)

async def process_guild_data(guild_id, paths):
    """Process data for a specific guild"""
    # Load players with better error handling
    players = {}
    try:
        with open(paths['character_file'], 'r', encoding='utf-8') as file:
            character_data = json.load(file)
            for playerInfo in character_data:
                player_name = playerInfo.get('name', '').strip()
                if player_name:
                    # Handle display_archetype - default to "DPS" if None or empty
                    display_archetype = playerInfo.get('display_archetype')
                    if display_archetype is None or not display_archetype.strip():
                        archetype = "DPS"
                    else:
                        archetype = display_archetype.strip()
                    
                    player = {
                        "name": player_name.capitalize(),
                        "archetype": archetype,
                    }
                    players[player_name.lower()] = player
        
        if not players:
            print("ℹ️ No players found in character file")
            return
            
        print(f"📋 Processing {len(players)} characters")
    except (json.JSONDecodeError, FileNotFoundError) as e:
        print(f"❌ Error loading character file: {e}")
        return
    
    try:
        # Check if required API credentials are available
        if not BLIZZARD_ID or not BLIZZARD_SECRET:
            error_msg = "Blizzard API credentials not available. Skipping armory update."
            print(f"❌ {error_msg}")
            logger.error(error_msg)
            return
        
        logger.info(f"Using Blizzard API credentials - ID: {BLIZZARD_ID[:8]}..., Secret: {'*' * len(BLIZZARD_SECRET)}")
        
        # Get Blizzard API token with retry logic
        access_token = None
        for attempt in range(3):
            try:
                logger.info(f"Attempting to get Blizzard API token (attempt {attempt + 1}/3)")
                response = requests.post(
                    BLIZZARD_TOKEN_URL, 
                    data={'grant_type': 'client_credentials'}, 
                    auth=(BLIZZARD_ID, BLIZZARD_SECRET),
                    timeout=10
                )
                response.raise_for_status()
                access_token = response.json()['access_token']
                logger.info("Successfully obtained Blizzard API token")
                break
            except requests.RequestException as e:
                error_msg = f"Token request attempt {attempt + 1} failed: {e}"
                print(f"⚠️ {error_msg}")
                logger.warning(error_msg)
                if attempt == 2:
                    final_error = "Failed to get Blizzard API token after 3 attempts"
                    print(f"❌ {final_error}")
                    logger.error(final_error)
                    return
                await asyncio.sleep(2)  # Wait before retry
        
        # Load existing armory data using guild-specific paths
        armory_data = {}
        if os.path.exists(paths['armory_file']):
            try:
                with open(paths['armory_file'], "r", encoding="utf-8") as f:
                    armory_data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                print("⚠️ Creating new armory file")
                armory_data = {}
        
        # Load existing parses data using guild-specific paths
        parses_data = {}
        if os.path.exists(paths['parses_file']):
            try:
                with open(paths['parses_file'], "r", encoding="utf-8") as f:
                    parses_data = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                print("⚠️ Creating new parses file")
                parses_data = {}
        
        # Get WCL API token with retry logic
        wcl_access_token = None
        if WCL_ID and WCL_SECRET:
            for attempt in range(3):
                try:
                    logger.info(f"Attempting to get WCL API token (attempt {attempt + 1}/3)")
                    wcl_token_url = "https://fresh.warcraftlogs.com/oauth/token"
                    wcl_data = {"grant_type": "client_credentials"}
                    wcl_response = requests.post(wcl_token_url, data=wcl_data, auth=(WCL_ID, WCL_SECRET), timeout=10)
                    wcl_response.raise_for_status()
                    wcl_access_token = wcl_response.json()["access_token"]
                    logger.info("Successfully obtained WCL API token")
                    break
                except requests.RequestException as e:
                    error_msg = f"WCL token request attempt {attempt + 1} failed: {e}"
                    print(f"⚠️ {error_msg}")
                    logger.warning(error_msg)
                    if attempt == 2:
                        final_error = "Failed to get WCL API token after 3 attempts"
                        print(f"❌ {final_error}")
                        logger.error(final_error)
                    await asyncio.sleep(2)  # Wait before retry
        else:
            logger.warning("WCL API credentials not available. Skipping WCL data updates.")
        
        # Fetch equipment and parses for each character with rate limiting
        new_items_found = 0
        new_parses_found = 0
        characters_processed = 0
        
        for player_key, player_info in players.items():
            try:
                # Rate limiting - small delay between API calls
                if characters_processed > 0:
                    await asyncio.sleep(0.7)  # 700ms delay between calls for both APIs
                
                player_name = player_info["name"]
                character_name = player_name.lower().replace(" ", "-")
                
                # Fetch WCL parses data (if WCL token is available)
                if wcl_access_token:
                    try:
                        wcl_server_slug = "spineshatter"
                        wcl_server_region = "EU"
                        wcl_api_url = "https://fresh.warcraftlogs.com/api/v2/client"
                        wcl_headers = {"Authorization": f"Bearer {wcl_access_token}"}
                        
                        # Determine metric based on archetype
                        metric = "hps" if player_info.get("archetype", "").lower() == "healer" else "bossdps"
                        
                        query = f"""
                        {{
                            characterData {{
                                character(name: "{player_key.lower()}", serverSlug: "{wcl_server_slug}", serverRegion: "{wcl_server_region}") {{
                                    zoneRankings(metric: {metric})
                                }}
                            }}
                        }}
                        """
                        
                        response = requests.post(wcl_api_url, headers=wcl_headers, json={"query": query}, timeout=10)
                        if response.status_code == 200:
                            wcl_data = response.json()
                            try:
                                # Extract zone rankings
                                character_data = wcl_data.get("data", {}).get("characterData", {})
                                character_info = character_data.get("character", {}) if character_data else {}
                                rankings = character_info.get("zoneRankings", {}) if character_info else {}
                                
                                # Parse rankings if it's a string
                                if isinstance(rankings, str):
                                    rankings = json.loads(rankings)
                                
                                # Create structured parse data
                                parse_info = {
                                    "metric": metric,
                                    "archetype": player_info.get("archetype", ""),
                                    "bestPerformanceAverage": rankings.get("bestPerformanceAverage", 0.0),
                                    "medianPerformanceAverage": rankings.get("medianPerformanceAverage", 0.0),
                                    "lastUpdated": datetime.now().isoformat()
                                }
                                
                                # Check if parses data changed
                                existing_parses = parses_data.get(player_name, {})
                                if (existing_parses.get("bestPerformanceAverage") != parse_info["bestPerformanceAverage"] or
                                    existing_parses.get("medianPerformanceAverage") != parse_info["medianPerformanceAverage"]):
                                    
                                    parses_data[player_name] = parse_info
                                    new_parses_found += 1
                                    print(f"📊 {player_name}: Updated parses (Best: {parse_info['bestPerformanceAverage']:.1f}, Median: {parse_info['medianPerformanceAverage']:.1f})")
                                
                            except (KeyError, json.JSONDecodeError, TypeError) as e:
                                print(f"⚠️ Error parsing WCL data for {player_name}: {e}")
                                # Set default values if parsing fails
                                parses_data[player_name] = {
                                    "metric": metric,
                                    "archetype": player_info.get("archetype", ""),
                                    "bestPerformanceAverage": 0.0,
                                    "medianPerformanceAverage": 0.0,
                                    "lastUpdated": datetime.now().isoformat(),
                                    "error": "Failed to parse WCL data"
                                }
                        elif response.status_code == 404:
                            print(f"⚠️ Character not found on WCL: {player_name}")
                        elif response.status_code == 429:
                            print(f"⚠️ WCL rate limited for {player_name}, waiting...")
                            await asyncio.sleep(5)
                        else:
                            print(f"⚠️ WCL API error for {player_name}: {response.status_code}")
                            
                    except Exception as e:
                        print(f"⚠️ Error fetching WCL parses for {player_name}: {str(e)[:100]}")
                
                # Fetch Blizzard armory data
                try:
                    url = f"https://eu.api.blizzard.com/profile/wow/character/spineshatter/{character_name}/equipment"
                    params = {
                        "namespace": "profile-classic1x-eu",
                        "locale": "en_GB"
                    }
                    headers = {'Authorization': f'Bearer {access_token}'}
                    
                    logger.debug(f"Fetching armory data for {player_name} from {url}")
                    
                    # Use async HTTP client for better performance
                    async with aiohttp.ClientSession() as session:
                        async with session.get(url, params=params, headers=headers, timeout=10) as response:
                            if response.status == 200:
                                data = await response.json()
                                equipped_items = [item["name"] for item in data.get("equipped_items", [])]
                                
                                # Initialize player in armory_data if not exists
                                if player_name not in armory_data:
                                    armory_data[player_name] = []
                                
                                # Check for new items
                                existing_items = set(armory_data[player_name])
                                new_items = [item for item in equipped_items if item not in existing_items]
                                
                                if new_items:
                                    armory_data[player_name].extend(new_items)
                                    new_items_found += len(new_items)
                                    print(f"🆕 {player_name}: {len(new_items)} new items")
                                    logger.info(f"Found {len(new_items)} new items for {player_name}")
                                    
                                    # Log each new item
                                    for item in new_items:
                                        logger.info(f"New item found for {player_name}: {item}")
                                else:
                                    logger.debug(f"No new items for {player_name}")
                            
                            elif response.status == 401:
                                error_msg = f"Unauthorized request for {player_name} - token may be expired"
                                print(f"❌ {error_msg}")
                                logger.error(error_msg)
                            elif response.status == 403:
                                error_msg = f"Forbidden request for {player_name} - check API credentials"
                                print(f"❌ {error_msg}")
                                logger.error(error_msg)
                            elif response.status == 404:
                                print(f"⚠️ Character not found on Blizzard API: {player_name}")
                                logger.warning(f"Character not found on Blizzard API: {player_name}")
                            elif response.status == 429:
                                print(f"⚠️ Blizzard API rate limited for {player_name}, waiting...")
                                logger.warning(f"Blizzard API rate limited for {player_name}")
                                await asyncio.sleep(5)
                            else:
                                error_msg = f"Blizzard API error for {player_name}: HTTP {response.status}"
                                print(f"⚠️ {error_msg}")
                                logger.error(error_msg)
                                # Log response text for debugging
                                try:
                                    response_text = await response.text()
                                    logger.error(f"Response body: {response_text[:200]}")
                                except:
                                    pass
                
                except aiohttp.ClientError as e:
                    error_msg = f"Network error fetching Blizzard armory for {player_name}: {str(e)}"
                    print(f"⚠️ {error_msg}")
                    logger.error(error_msg)
                except Exception as e:
                    error_msg = f"Unexpected error fetching Blizzard armory for {player_name}: {str(e)}"
                    print(f"⚠️ {error_msg}")
                    logger.error(error_msg, exc_info=True)
                
                characters_processed += 1
                
            except asyncio.TimeoutError:
                print(f"⚠️ Timeout fetching data for {player_name}")
            except Exception as e:
                print(f"⚠️ Error processing {player_name}: {str(e)[:100]}")
        
        # Save updated armory data atomically
        if new_items_found > 0 or characters_processed > 0:
            try:
                # Sort armory data alphabetically by character name before saving
                sorted_armory_data = dict(sorted(armory_data.items()))
                
                # Write to temporary file first, then rename (atomic operation)
                temp_file = paths['armory_file'] + '.tmp'
                with open(temp_file, "w", encoding="utf-8") as f:
                    json.dump(sorted_armory_data, f, ensure_ascii=False, indent=2)
                
                # Atomic rename
                os.replace(temp_file, paths['armory_file'])
                print(f"💾 Armory data saved - {new_items_found} new items found")
            except Exception as e:
                print(f"❌ Error saving armory data: {e}")
                # Clean up temp file if it exists
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        
        # Save updated parses data atomically
        if new_parses_found > 0 or characters_processed > 0:
            try:
                # Write to temporary file first, then rename (atomic operation)
                temp_file = paths['parses_file'] + '.tmp'
                with open(temp_file, "w", encoding="utf-8") as f:
                    json.dump(parses_data, f, ensure_ascii=False, indent=2)
                
                # Atomic rename
                os.replace(temp_file, paths['parses_file'])
                print(f"� Parses data saved - {new_parses_found} characters updated")
            except Exception as e:
                print(f"❌ Error saving parses data: {e}")
                # Clean up temp file if it exists
                if os.path.exists(temp_file):
                    os.remove(temp_file)
        
        print(f"✅ Data update completed - {characters_processed}/{len(players)} characters processed")
        logger.info(f"Data update completed - {characters_processed}/{len(players)} characters processed")
        print(f"   📊 Summary: {new_items_found} new items, {new_parses_found} parse updates")
        logger.info(f"Summary: {new_items_found} new items, {new_parses_found} parse updates")
        
    except Exception as e:
        print(f"❌ Critical error in periodic task: {e}")
        logger.error(f"Critical error in periodic task: {e}", exc_info=True)

@periodic_task.before_loop
async def before_periodic_task():
    await bot.wait_until_ready()
    print("🚀 Periodic task started - will run every 5 minutes")
    logger.info("Periodic task started - will run every 5 minutes")

async def validate_character_exists(character_name):
    try:
        url = f"https://classicwowarmory.com/character/eu/spineshatter/{character_name.replace(' ', '%20')}"
        
        # Check if the character exists
        async with aiohttp.ClientSession() as session:
            async with session.get(url, timeout=10) as response:
                if "/character/404" in str(response.url) or response.status != 200:
                    return False, "Character not found on armory"
                return True, None
                    
    except Exception as e:
        return False, f"Error checking character: {str(e)}"

async def get_staff_mentions(guild):
    mentions = []
    
    # Look for Karumenta
    karumenta = discord.utils.get(guild.members, name="Karumenta")
    if karumenta:
        mentions.append(karumenta.mention)
    else:
        mentions.append("**Karumenta**")
    
    # Look for Hokkies
    hokkies = discord.utils.get(guild.members, name="Hokkies")
    if hokkies:
        mentions.append(hokkies.mention)
    else:
        mentions.append("**Hokkies**")
    
    return " or ".join(mentions)

async def validate_character_name(character_name, guild=None):
    try:
        url = f"https://classicwowarmory.com/character/eu/spineshatter/{character_name.replace(' ', '%20')}"
        response = requests.get(url, timeout=10, allow_redirects=True)
        print(url)
        
        # Get staff mentions
        staff_mentions = "**Karumenta** or **Hokkies**"
        if guild:
            staff_mentions = await get_staff_mentions(guild)
        
        # Check if we were redirected to the 404 page
        if "/character/404" in response.url or response.status_code == 404:
            return False, f"Character not found on Spineshatter realm. Please check the spelling and try again. If it's an error, please contact {staff_mentions} for assistance."
        elif response.status_code == 200:
            return True, None
        else:
            return False, f"Unable to verify character (HTTP {response.status_code}). Please contact {staff_mentions} for assistance."
    
    except requests.exceptions.Timeout:
        staff_mentions = "**Karumenta** or **Hokkies**"
        if guild:
            staff_mentions = await get_staff_mentions(guild)
        return False, f"Connection timeout while verifying character. Please contact {staff_mentions} for assistance."
    except requests.exceptions.RequestException as e:
        staff_mentions = "**Karumenta** or **Hokkies**"
        if guild:
            staff_mentions = await get_staff_mentions(guild)
        return False, f"Network error while verifying character. Please contact {staff_mentions} for assistance."
    except Exception as e:
        staff_mentions = "**Karumenta** or **Hokkies**"
        if guild:
            staff_mentions = await get_staff_mentions(guild)
        return False, f"Unexpected error while verifying character. Please contact {staff_mentions} for assistance."

class ApplicationView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)  # Persistent view
    
    @discord.ui.button(label='Apply', style=discord.ButtonStyle.green, emoji='📝')
    async def apply_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        user = interaction.user
        
        # Check if user already has Trial, Raider, Officer, or Guild Leader role
        guild = interaction.guild
        member = guild.get_member(user.id)
        if member:
            trial_role = discord.utils.get(guild.roles, name="Trial")
            raider_role = discord.utils.get(guild.roles, name="Raider")
            officer_role = discord.utils.get(guild.roles, name="Officer")
            guild_leader_role = discord.utils.get(guild.roles, name="Guild Leader")
            
            # Get staff mentions
            staff_mentions = await get_staff_mentions(guild)
            
            if trial_role and trial_role in member.roles:
                await interaction.response.send_message(f"❌ You already have the **Trial** role and cannot apply again. If you need assistance, please contact {staff_mentions}.", ephemeral=True)
                return
            
            if raider_role and raider_role in member.roles:
                await interaction.response.send_message(f"❌ You already have the **Raider** role and cannot apply again. If you need assistance, please contact {staff_mentions}.", ephemeral=True)
                return
            
            if officer_role and officer_role in member.roles:
                await interaction.response.send_message(f"❌ You already have the **Officer** role and cannot apply again. If you need assistance, please contact {staff_mentions}.", ephemeral=True)
                return
            
            if guild_leader_role and guild_leader_role in member.roles:
                await interaction.response.send_message(f"❌ You already have the **Guild Leader** role and cannot apply again. If you need assistance, please contact {staff_mentions}.", ephemeral=True)
                return
        
        # Check if user already has an active application
        if user.id in active_applications:
            await interaction.response.send_message("❌ You already have an active application in progress. Please complete it first or wait for it to expire.", ephemeral=True)
            return
        
        try:
            # Initialize application data
            active_applications[user.id] = {
                'question_index': 0,
                'answers': [],
                'guild_id': interaction.guild.id,
                'start_time': asyncio.get_event_loop().time()  # Track when application started
            }
            
            # Send first question
            embed = discord.Embed(
                title="🎉 Application Started!",
                description=f"Thank you for your interest in applying! I'll ask you some questions.",
                color=0x00ff00
            )
            embed.add_field(
                name=f"Question 1/{len(APPLICATION_QUESTIONS)}",
                value=APPLICATION_QUESTIONS[0],
                inline=False
            )
            embed.set_footer(text="Please respond with your answer. Type 'cancel' to cancel the application.")
            
            await user.send(embed=embed)
            
            # Respond to the interaction
            await interaction.response.send_message("✅ Check your DMs! I've started your application process.", ephemeral=True)
            
        except discord.Forbidden:
            # User has DMs disabled
            await interaction.response.send_message("❌ I couldn't send you a DM. Please enable DMs from server members and try again.", ephemeral=True)
        except Exception as e:
            await interaction.response.send_message("❌ An error occurred. Please try again later or reach someone from the Staff.", ephemeral=True)
            print(f"Error sending DM: {e}")

class ReviewView(discord.ui.View):
    def __init__(self, user_id, character_name, application_channel, review_channel):
        super().__init__(timeout=None)
        self.user_id = user_id
        self.character_name = character_name
        self.application_channel = application_channel
        self.review_channel = review_channel
    
    @discord.ui.button(label='Accept', style=discord.ButtonStyle.green, emoji='✅')
    async def accept_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        guild = interaction.guild
        member = guild.get_member(self.user_id)
        
        if not member:
            await interaction.response.send_message("❌ User not found in server.", ephemeral=True)
            return
        
        try:
            # Get or create "Trial" role
            trial_role = discord.utils.get(guild.roles, name="Trial")
            if not trial_role:
                trial_role = await guild.create_role(name="Trial")
                print("Created 'Trial' role!")
            
            # Remove "Social" role if user has it
            social_role = discord.utils.get(guild.roles, name="Social")
            if social_role and social_role in member.roles:
                await member.remove_roles(social_role)
                print(f"Removed 'Social' role from {member.display_name}")
            
            # Give user the Trial role
            await member.add_roles(trial_role)
            
            # Get or create "Trials" category with proper permissions
            trials_category = discord.utils.get(guild.categories, name="Trials")
            if not trials_category:
                # Set up permissions for Trials category (Officer, Bot, Guild Leader only)
                overwrites = {
                    guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False),
                }
                
                # Add permissions for specific roles
                officer_role = discord.utils.get(guild.roles, name="Officer")
                guild_leader_role = discord.utils.get(guild.roles, name="Guild Leader")
                bot_member = guild.me  # The bot itself
                
                if officer_role:
                    overwrites[officer_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
                if guild_leader_role:
                    overwrites[guild_leader_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
                if bot_member:
                    overwrites[bot_member] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
                
                trials_category = await guild.create_category("Trials", overwrites=overwrites)
                print("Created 'Trials' category with Officer/Guild Leader/Bot permissions!")
            
            # Rename application channel and move both channels to Trials category
            if self.application_channel:
                new_channel_name = f"trial-{self.character_name.lower().replace(' ', '-')}"
                # When moving to Trials, keep user access for the trial channel
                trial_overwrites = {
                    guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False),
                    member: discord.PermissionOverwrite(read_messages=True, send_messages=True, view_channel=True)
                }
                await self.application_channel.edit(name=new_channel_name, category=trials_category, overwrites=trial_overwrites)
            if self.review_channel:
                # Review channel inherits category permissions only
                await self.review_channel.edit(category=trials_category)
            
            # Send acceptance message to application channel
            if self.application_channel:
                accept_embed = discord.Embed(
                    title="🎉 Application Accepted!",
                    description=f"Congratulations {member.mention}! Your application has been accepted and you've been given the **Trial** role.",
                    color=0x00ff00
                )
                accept_embed.add_field(
                    name="📋 General Information",
                    value="Just some general info we work on a no sign up based roster, post in ⁠⛔absence if you're going to miss a raid, so i won't roster you for that week, and i try to post the roster around friday in ⁠📒raid-assigments and update the assignments with it.",
                    inline=False
                )
                accept_embed.add_field(
                    name="🎯 TMB Setup Required",
                    value="Please create a character on https://thatsmybis.com/ and add him to the guild from the home page. Once you do it notify us, Thanks!",
                    inline=False
                )
                accept_embed.add_field(
                    name="⚙️ Addons Required",
                    value="Please make sure you install RCLC lootcouncil before heading into your first raid with us, we use this addon to distribute loot in our raids 🙂",
                    inline=False
                )
                await self.application_channel.send(embed=accept_embed)
            
            # Send confirmation message to review channel
            await interaction.response.send_message(f"✅ Application accepted by {interaction.user.mention}! {member.mention} has been given the Trial role and channels moved to Trials category.", ephemeral=False)
            
            # Find and delete the first message (review message with buttons)
            if self.review_channel:
                async for message in self.review_channel.history(limit=50, oldest_first=True):
                    if message.author == bot.user and message.embeds:
                        # Check if this is the review message by looking for the title
                        embed = message.embeds[0]
                        if embed.title and "Review Application" in embed.title:
                            await message.delete()
                            break
            
        except Exception as e:
            await interaction.response.send_message(f"❌ Error processing acceptance: {e}", ephemeral=True)
            print(f"Error accepting application: {e}")
    
    @discord.ui.button(label='Decline', style=discord.ButtonStyle.red, emoji='❌')
    async def decline_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        try:
            # Send decline message to application channel before deletion
            if self.application_channel:
                decline_embed = discord.Embed(
                    title="❌ Application Declined",
                    description=f"Unfortunately {self.character_name}, your application has been declined. You may reapply in the future.",
                    color=0xff0000
                )
                await self.application_channel.send(embed=decline_embed)
            
            await interaction.response.send_message("❌ Application declined. Review channel will be deleted.", ephemeral=False)
            
            # Delete the review channel
            if self.review_channel:
                await self.review_channel.delete()
            
        except Exception as e:
            await interaction.response.send_message(f"❌ Error processing decline: {e}", ephemeral=True)
            print(f"Error declining application: {e}")

@bot.event
async def on_message(message):
    # Ignore bot messages
    if message.author.bot:
        return
    
    # Check if this is a DM and user has an active application
    if isinstance(message.channel, discord.DMChannel) and message.author.id in active_applications:
        await handle_application_response(message)
        return
    
    # Process commands
    await bot.process_commands(message)

async def handle_application_response(message):
    user_id = message.author.id
    app_data = active_applications[user_id]
    
    # Check for cancel command
    if message.content.lower() == 'cancel':
        del active_applications[user_id]
        embed = discord.Embed(
            title="❌ Application Cancelled",
            description="Your application has been cancelled. You can start a new one anytime by clicking the Apply button again.",
            color=0xff0000
        )
        await message.channel.send(embed=embed)
        return
    
    # Special validation for character name (first question) - COMMENTED OUT
    # if app_data['question_index'] == 0:
    #     character_name = message.content.strip()
    #     
    #     # Validate character name
    #     guild = bot.get_guild(app_data['guild_id'])
    #     is_valid, error_message = await validate_character_name(character_name, guild)
    #     
    #     if not is_valid:
    #         # Send error message and ask for character name again
    #         error_embed = discord.Embed(
    #             title="❌ Character Validation Failed",
    #             description=error_message,
    #             color=0xff0000
    #         )
    #         error_embed.add_field(
    #             name=f"Question 1/{len(APPLICATION_QUESTIONS)}",
    #             value=APPLICATION_QUESTIONS[0],
    #             inline=False
    #         )
    #         error_embed.set_footer(text="Please provide the correct character name or type 'cancel' to cancel the application.")
    #         await message.channel.send(embed=error_embed)
    #         return  # Don't advance to next question, ask again
    
    # Save the answer
    app_data['answers'].append(message.content)
    app_data['question_index'] += 1
    
    # Check if we have more questions
    if app_data['question_index'] < len(APPLICATION_QUESTIONS):
        # Send next question
        embed = discord.Embed(
            title="📝 Next Question",
            color=0x00ff00
        )
        embed.add_field(
            name=f"Question {app_data['question_index'] + 1}/{len(APPLICATION_QUESTIONS)}",
            value=APPLICATION_QUESTIONS[app_data['question_index']],
            inline=False
        )
        embed.set_footer(text="Please respond with your answer. Type 'cancel' to cancel the application.")
        
        await message.channel.send(embed=embed)
    else:
        # Application completed
        await complete_application(message.author, app_data)

async def complete_application(user, app_data):
    # Remove from active applications
    del active_applications[user.id]
    
    # Get the guild
    guild = bot.get_guild(app_data['guild_id'])
    if not guild:
        return
    
    # Get the user's member object in the guild
    member = guild.get_member(user.id)
    if not member:
        return
    
    # Get the character name from the first answer
    character_name = app_data['answers'][0] if app_data['answers'] else "Unknown"
    
    # Store the old nickname for comparison
    old_nick = member.display_name
    
    # Rename the user to their character name
    try:
        await member.edit(nick=character_name)
        print(f"Renamed {user.display_name} to {character_name}")
    except discord.Forbidden:
        print(f"Cannot rename {user.display_name} - insufficient permissions")
    except Exception as e:
        print(f"Error renaming user: {e}")
    
    # Check if "Applications" category exists
    applications_category = discord.utils.get(guild.categories, name="Applications")
    
    if not applications_category:
        # Set up permissions for Applications category (Officer, Bot, Guild Leader only)
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False),
        }
        
        # Add permissions for specific roles
        officer_role = discord.utils.get(guild.roles, name="Officer")
        guild_leader_role = discord.utils.get(guild.roles, name="Guild Leader")
        bot_member = guild.me  # The bot itself
        
        if officer_role:
            overwrites[officer_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if guild_leader_role:
            overwrites[guild_leader_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if bot_member:
            overwrites[bot_member] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        
        # Create the "Applications" category with permissions
        applications_category = await guild.create_category("Applications", overwrites=overwrites)
        print("Created 'Applications' category with Officer/Guild Leader/Bot permissions!")
    
    # Create application channel with user access (inherits category permissions + user access)
    application_channel_name = f"application-{character_name.lower().replace(' ', '-')}"
    try:
        # Set permissions for the application channel (inherits from category + user access)
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False),
            member: discord.PermissionOverwrite(read_messages=True, send_messages=True, view_channel=True)
        }
        
        application_channel = await guild.create_text_channel(
            application_channel_name,
            category=applications_category,
            overwrites=overwrites
        )
        print(f"Created application channel: {application_channel_name}")
    except Exception as e:
        print(f"Error creating application channel: {e}")
        application_channel = None
    
    # Create review channel (inherits category permissions only)
    review_channel_name = f"review-{character_name.lower().replace(' ', '-')}"
    try:
        # Set explicit permissions for review channel (deny @everyone, inherit officer/guild leader from category)
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        
        review_channel = await guild.create_text_channel(
            review_channel_name,
            category=applications_category,
            overwrites=overwrites
        )
        print(f"Created review channel: {review_channel_name}")
    except Exception as e:
        print(f"Error creating review channel: {e}")
        review_channel = None
    
    # Send completion message to user
    embed = discord.Embed(
        title="🎉 Application Completed!",
        description=f"Thank you for completing your application, {character_name}! Our staff will review it and get back to you soon.",
        color=0x00ff00
    )
    if application_channel:
        embed.add_field(
            name="Your Application Channel",
            value=f"You can check the status of your application in {application_channel.mention}",
            inline=False
        )
    await user.send(embed=embed)
    
    # Post the application in the application channel
    if application_channel:
        embed = discord.Embed(
            title=f"📋 Application for {character_name}",
            description=f"Application submitted by {user.mention}",
            color=0x0099ff
        )
        
        for i, (question, answer) in enumerate(zip(APPLICATION_QUESTIONS, app_data['answers'])):
            embed.add_field(
                name=f"Q{i+1}: {question}",
                value=answer[:1024] if len(answer) <= 1024 else answer[:1021] + "...",
                inline=False
            )
        
        embed.set_footer(text=f"User ID: {user.id}")
        await application_channel.send(embed=embed)
        
        # Send nick change notification if nickname was different
        if old_nick != character_name:
            nick_embed = discord.Embed(
                title="📝 Nickname Updated",
                description=f"Your server nickname has been changed from **{old_nick}** to **{character_name}**",
                color=0x00ff00
            )
            await application_channel.send(embed=nick_embed)
    
    # Send review message with Accept/Decline buttons to review channel for staff
    if review_channel:
        embed = discord.Embed(
            title=f"📋 Review Application - {character_name}",
            description=f"**Applicant:** {user.mention} ({user.display_name})\n**Character Name:** {character_name}",
            color=0xffa500
        )
        embed.add_field(
            name="📁 Application Details",
            value=f"Full application can be viewed in {application_channel.mention if application_channel else 'application channel'}",
            inline=False
        )
        embed.add_field(
            name="⚡ Actions",
            value="Click **Accept** to give Trial role and move to Trials category\nClick **Decline** to reject and delete this review channel",
            inline=False
        )
        
        view = ReviewView(user.id, character_name, application_channel, review_channel)
        await review_channel.send(embed=embed, view=view)
        
        # Send character lookup links
        links_embed = discord.Embed(
            title="🔗 Character Lookup Links",
            color=0x0099ff
        )
        links_embed.add_field(
            name="Warcraft Logs",
            value=f"[View WCL Profile](https://fresh.warcraftlogs.com/character/eu/spineshatter/{character_name.replace(' ', '%20')})",
            inline=False
        )
        links_embed.add_field(
            name="Classic WoW Armory",
            value=f"[View Armory Profile](https://classicwowarmory.com/character/eu/spineshatter/{character_name.replace(' ', '%20')})",
            inline=False
        )
        
        # Check if character exists and add a note if not found
        character_exists, error_msg = await validate_character_exists(character_name)
        if not character_exists and error_msg:
            links_embed.add_field(
                name="⚠️ Note",
                value=f"Character validation: {error_msg}",
                inline=False
            )
        
        await review_channel.send(embed=links_embed)

@bot.event
@bot.event
async def on_ready():
    print(f'Logged in as {bot.user.name} - {bot.user.id}')
    logger.info(f'Bot logged in as {bot.user.name} - {bot.user.id}')
    print(f'Bot is in {len(bot.guilds)} guilds')
    logger.info(f'Bot is in {len(bot.guilds)} guilds')
    print('------')
    
    # Start the periodic task
    if not periodic_task.is_running():
        periodic_task.start()
        logger.info('Periodic task started')

@bot.event
async def on_error(event, *args, **kwargs):
    """Handle general bot errors"""
    print(f'Error in event {event}: {args}')

@bot.event  
async def on_command_error(ctx, error):
    """Handle command errors"""
    if isinstance(error, commands.CommandNotFound):
        return  # Ignore unknown commands
    elif isinstance(error, commands.MissingPermissions):
        await ctx.send("❌ You don't have permission to use this command.")
    else:
        await ctx.send(f"❌ An error occurred: {str(error)}")
        print(f'Command error: {error}')

class BotManagementView(discord.ui.View):
    def __init__(self):
        super().__init__(timeout=None)  # Persistent view
    
    @discord.ui.button(label='Get Attendance', style=discord.ButtonStyle.primary, emoji='📊')
    async def get_attendance_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        # Check permissions - only allow Officers, Guild Leaders
        authorized_roles = ["Officer", "Guild Leader"]
        user_roles = [role.name for role in interaction.user.roles]
        
        if not any(role in authorized_roles for role in user_roles):
            await interaction.response.send_message("❌ You don't have permission to use this feature. Required roles: Officer or Guild Leader.", ephemeral=True)
            return
        
        try:
            # Get guild-specific data
            guild_id = interaction.guild.id
            paths = get_guild_file_paths(guild_id)
            
            # Initialize guild data if needed
            initialize_guild_data_files(guild_id)
            
            # Generate Excel for this guild
            await interaction.response.send_message("📊 Generating attendance report for this server...", ephemeral=True)
            
            # Create guild-specific Excel (Attendance type)
            workbook = createExcel(guild_id, "Attendance")
            
            if workbook:
                # Save and send the file
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"attendance_sheet_{interaction.guild.name}_{timestamp}.xlsx"
                file_path = os.path.join(paths['sheet_dir'], filename)
                
                workbook.save(file_path)
                
                # Edit the original message to send the file
                await interaction.edit_original_response(
                    content=f"📊 **Attendance Report for {interaction.guild.name}**\nGenerated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\nThis report contains attendance data for all guild members.",
                    attachments=[discord.File(file_path, filename=filename)]
                )
                
                # Clean up file
                os.remove(file_path)
            else:
                await interaction.edit_original_response(content="❌ Failed to generate Excel file. Please check the data files.")
            
        except Exception as e:
            await interaction.edit_original_response(content=f"❌ Error generating Excel report: {str(e)[:100]}...")
            logger.error(f"Error in attendance button: {e}", exc_info=True)
    
    @discord.ui.button(label='Get Class Items', style=discord.ButtonStyle.primary, emoji='⚔️')
    async def get_class_items_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        # Check permissions - only allow Officers, Guild Leaders
        authorized_roles = ["Officer", "Guild Leader"]
        user_roles = [role.name for role in interaction.user.roles]
        
        if not any(role in authorized_roles for role in user_roles):
            await interaction.response.send_message("❌ You don't have permission to use this feature. Required roles: Officer or Guild Leader.", ephemeral=True)
            return
        
        try:
            # Get guild-specific data
            guild_id = interaction.guild.id
            paths = get_guild_file_paths(guild_id)
            
            # Initialize guild data if needed
            initialize_guild_data_files(guild_id)
            
            # Generate Excel for this guild
            await interaction.response.send_message("⚔️ Generating class items report for this server...", ephemeral=True)
            
            # Create guild-specific Excel (Class Items type)
            workbook = createExcel(guild_id, "Class Items")
            
            if workbook:
                # Save and send the file
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"class_items_sheet_{interaction.guild.name}_{timestamp}.xlsx"
                file_path = os.path.join(paths['sheet_dir'], filename)
                
                workbook.save(file_path)
                
                # Edit the original message to send the file
                await interaction.edit_original_response(
                    content=f"⚔️ **Class Items Report for {interaction.guild.name}**\nGenerated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\nThis report contains class-specific item data and recommendations.",
                    attachments=[discord.File(file_path, filename=filename)]
                )
                
                # Clean up file
                os.remove(file_path)
            else:
                await interaction.edit_original_response(content="❌ Failed to generate Excel file. Please check the data files.")
            
        except Exception as e:
            await interaction.edit_original_response(content=f"❌ Error generating class items report: {str(e)[:100]}...")
            logger.error(f"Error in class items button: {e}", exc_info=True)
    
    @discord.ui.button(label='Get Loot', style=discord.ButtonStyle.primary, emoji='💎')
    async def get_loot_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        # Check permissions - only allow Officers, Guild Leaders
        authorized_roles = ["Officer", "Guild Leader"]
        user_roles = [role.name for role in interaction.user.roles]
        
        if not any(role in authorized_roles for role in user_roles):
            await interaction.response.send_message("❌ You don't have permission to use this feature. Required roles: Officer or Guild Leader.", ephemeral=True)
            return
        
        try:
            # Get guild-specific data
            guild_id = interaction.guild.id
            paths = get_guild_file_paths(guild_id)
            
            # Initialize guild data if needed
            initialize_guild_data_files(guild_id)
            
            # Generate Excel for this guild
            await interaction.response.send_message("💎 Generating loot distribution report for this server...", ephemeral=True)
            
            # Create guild-specific Excel (Loot type)
            workbook = createExcel(guild_id, "Loot")
            
            if workbook:
                # Save and send the file
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"loot_template_sheet_{interaction.guild.name}_{timestamp}.xlsx"
                file_path = os.path.join(paths['sheet_dir'], filename)
                
                workbook.save(file_path)
                
                # Edit the original message to send the file
                await interaction.edit_original_response(
                    content=f"💎 **Loot Distribution Report for {interaction.guild.name}**\nGenerated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\nThis report contains guild loot distribution data and analytics.",
                    attachments=[discord.File(file_path, filename=filename)]
                )
                
                # Clean up file
                os.remove(file_path)
            else:
                await interaction.edit_original_response(content="❌ Failed to generate Excel file. Please check the data files.")
            
        except Exception as e:
            await interaction.edit_original_response(content=f"❌ Error generating loot report: {str(e)[:100]}...")
            logger.error(f"Error in loot button: {e}", exc_info=True)
    
    @discord.ui.button(label='Get All', style=discord.ButtonStyle.success, emoji='📦')
    async def get_all_button(self, interaction: discord.Interaction, button: discord.ui.Button):
        # Check permissions - only allow Officers, Guild Leaders
        authorized_roles = ["Officer", "Guild Leader"]
        user_roles = [role.name for role in interaction.user.roles]
        
        if not any(role in authorized_roles for role in user_roles):
            await interaction.response.send_message("❌ You don't have permission to use this feature. Required roles: Officer or Guild Leader.", ephemeral=True)
            return
        
        # This button does the same as Get Attendance for now (generates the main Excel file)
        try:
            # Get guild-specific data
            guild_id = interaction.guild.id
            paths = get_guild_file_paths(guild_id)
            
            # Initialize guild data if needed
            initialize_guild_data_files(guild_id)
            
            # Generate Excel for this guild
            await interaction.response.send_message("📦 Generating comprehensive guild data export for this server...", ephemeral=True)
            
            # Create guild-specific Excel (All types)
            workbook = createExcel(guild_id, "All")
            
            if workbook:
                # Save and send the file
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"complete_guild_export_{interaction.guild.name}_{timestamp}.xlsx"
                file_path = os.path.join(paths['sheet_dir'], filename)
                
                workbook.save(file_path)
                
                # Edit the original message to send the file
                await interaction.edit_original_response(
                    content=f"📦 **Complete Guild Data Export for {interaction.guild.name}**\nGenerated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\nThis file contains all available guild data including attendance, armory items, and performance statistics.",
                    attachments=[discord.File(file_path, filename=filename)]
                )
                
                # Clean up file
                os.remove(file_path)
            else:
                await interaction.edit_original_response(content="❌ Failed to generate Excel file. Please check the data files.")
            
        except Exception as e:
            await interaction.edit_original_response(content=f"❌ Error generating comprehensive export: {str(e)[:100]}...")
            logger.error(f"Error in get all button: {e}", exc_info=True)

@bot.command()
async def setupHopium(ctx):
    guild = ctx.guild
    setup_messages = []  # Track messages to delete later
    
    # Check if "Recruitment" category exists
    recruitment_category = discord.utils.get(guild.categories, name="Recruitment")
    
    if not recruitment_category:
        # Create the "Recruitment" category with restricted permissions
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        recruitment_category = await guild.create_category("Recruitment", overwrites=overwrites)
        msg = await ctx.send("Created 'Recruitment' category!")
        setup_messages.append(msg)
    
    # Check if "apply-here" channel exists in the category
    apply_channel = discord.utils.get(recruitment_category.channels, name="✍apply-here")
    
    if not apply_channel:
        # Create the "apply-here" channel in the Recruitment category with restricted permissions
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        apply_channel = await guild.create_text_channel("✍apply-here", category=recruitment_category, overwrites=overwrites)
        msg = await ctx.send("Created 'apply-here' channel in the Recruitment category!")
        setup_messages.append(msg)
    
    # Clear the apply-here channel to ensure clean state
    await apply_channel.purge()
    msg = await ctx.send("Cleared 'apply-here' channel!")
    setup_messages.append(msg)
    
    # Create the application message with button
    embed = discord.Embed(
        title="📋 Application for Hopium Guild",
        description=f"Click the button below to start your application process!\nIf anything goes wrong, please contact {await get_staff_mentions(guild)}.",
        color=0x00ff00
    )
    
    view = ApplicationView()
    await apply_channel.send(embed=embed, view=view)
    msg = await ctx.send("Application message sent with Apply button!")
    setup_messages.append(msg)
    
    # Check if "ADMIN" category exists
    admin_category = discord.utils.get(guild.categories, name="ADMIN")
    
    if not admin_category:
        # Create the "ADMIN" category with restricted permissions (Officers and Guild Leaders only)
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        
        # Add permissions for specific roles
        officer_role = discord.utils.get(guild.roles, name="Officer")
        guild_leader_role = discord.utils.get(guild.roles, name="Guild Leader")
        bot_member = guild.me  # The bot itself
        
        if officer_role:
            overwrites[officer_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if guild_leader_role:
            overwrites[guild_leader_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if bot_member:
            overwrites[bot_member] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        
        admin_category = await guild.create_category("ADMIN", overwrites=overwrites)
        msg = await ctx.send("Created 'ADMIN' category!")
        setup_messages.append(msg)
    
    # Check if "HopiumBot" channel exists in the ADMIN category
    hopium_bot_channel = discord.utils.get(admin_category.channels, name="🤖hopiumbot")
    
    if not hopium_bot_channel:
        # Create the "HopiumBot" channel in the ADMIN category (inherits category permissions)
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        
        # Add permissions for specific roles
        officer_role = discord.utils.get(guild.roles, name="Officer")
        guild_leader_role = discord.utils.get(guild.roles, name="Guild Leader")
        bot_member = guild.me
        
        if officer_role:
            overwrites[officer_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if guild_leader_role:
            overwrites[guild_leader_role] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        if bot_member:
            overwrites[bot_member] = discord.PermissionOverwrite(read_messages=True, send_messages=True, manage_messages=True, view_channel=True)
        
        hopium_bot_channel = await guild.create_text_channel("🤖hopiumbot", category=admin_category, overwrites=overwrites)
        msg = await ctx.send("Created 'HopiumBot' channel in the ADMIN category!")
        setup_messages.append(msg)
    
    # Clear the HopiumBot channel to ensure clean state
    await hopium_bot_channel.purge()
    msg = await ctx.send("Cleared 'HopiumBot' channel!")
    setup_messages.append(msg)
    
    # Create the bot guide message with management buttons
    guide_embed = discord.Embed(
        title="🤖 HopiumBot Management Panel",
        description="Welcome to the HopiumBot management interface! Use the buttons below for quick access to guild data or the commands listed for advanced operations.",
        color=0x9932cc
    )
    
    guide_embed.add_field(
        name="📋 File Download Commands",
        value="• `!getfile armory` - Download guild armory data\n• `!getfile icons` - Download item icons data\n• `!getfile parses` - Download guild WCL parses data\n• `!getfile tmb` - Download TMB files (character, attendance, item notes)",
        inline=False
    )
    
    guide_embed.add_field(
        name="👤 Player Data Commands", 
        value="• `!get armory <playerName>` - Get specific player's armory data\n• `!get parses <playerName>` - Get specific player's WCL parses",
        inline=False
    )
    
    guide_embed.add_field(
        name="📤 Upload Commands",
        value="• `!uploadtmb` - Upload TMB files (character-json.json, hopium-attendance.csv, item-notes.csv)\n• `!uploadarmory` - Upload armory.json file (merges with existing data)",
        inline=False
    )
    
    guide_embed.add_field(
        name="⚙️ Management Commands",
        value="• `!setupHopium` - Run initial bot setup (creates channels and categories)\n• All commands are restricted to Officers and Guild Leaders only",
        inline=False
    )
    
    guide_embed.add_field(
        name="📊 Excel Generation",
        value="Use the buttons below to generate Excel reports:\n• **Get Attendance** - Guild attendance sheet\n• **Get Class Items** - Class-specific item sheet\n• **Get Loot** - Loot prio template sheet\n• **Get All** - All of the above in one file",
        inline=False
    )
    
    guide_embed.set_footer(text="Click the buttons below for Excel generation features • Officers & Guild Leaders only")
    
    management_view = BotManagementView()
    await hopium_bot_channel.send(embed=guide_embed, view=management_view)
    msg = await ctx.send("Bot management panel created with guide and buttons!")
    setup_messages.append(msg)
    
    msg = await ctx.send('✅ Setup completed! Messages will be deleted in 5 seconds...')
    setup_messages.append(msg)
    
    # Wait 5 seconds then delete all setup messages including the command message
    await asyncio.sleep(5)
    
    # Delete the original command message
    try:
        await ctx.message.delete()
    except discord.NotFound:
        pass
    
    # Delete all setup status messages
    for message in setup_messages:
        try:
            await message.delete()
        except discord.NotFound:
            pass

@bot.command(name='getfile')
async def get_file_data(ctx, data_type: str = None):
    """
    Get guild data files as temporary messages
    Usage: !getfile armory | !getfile icons
    Restricted to Officers, Guild Leaders, and authorized roles
    """
    # Check permissions - only allow Officers, Guild Leaders
    authorized_roles = ["Officer", "Guild Leader"]
    user_roles = [role.name for role in ctx.author.roles]
    
    if not any(role in authorized_roles for role in user_roles):
        await ctx.send("❌ You don't have permission to use this command. Required roles: Officer or Guild Leader.", delete_after=10)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    if data_type not in ['armory', 'icons', 'tmb', 'parses']:
        embed = discord.Embed(
            title="📋 Available Data Types",
            description=f"Choose from available data file types for **{ctx.guild.name}**:",
            color=0xff9900
        )
        embed.add_field(
            name="📥 Download Commands",
            value="`!getfile armory` - Download guild armory data\n`!getfile icons` - Download item icons data\n`!getfile parses` - Download guild WCL parses data\n`!getfile tmb` - Download TMB data files (character, attendance, item notes)",
            inline=False
        )
        embed.add_field(
            name="📤 Upload Commands",
            value="`!uploadtmb` - Upload TMB files (character-json.json, hopium-attendance.csv, item-notes.csv)\n`!uploadarmory` - Upload armory.json file (merges with existing data)",
            inline=False
        )
        embed.add_field(
            name="🏰 Server Info",
            value=f"**Guild:** {ctx.guild.name}\n**All data operations are server-specific**",
            inline=False
        )
        embed.set_footer(text=f"Guild: {ctx.guild.name} • Data is server-specific")
        await ctx.send(embed=embed, delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Get guild-specific file paths
    guild_id = ctx.guild.id
    paths = get_guild_file_paths(guild_id)
    
    # Initialize guild data if needed
    initialize_guild_data_files(guild_id)
    
    try:
        # Handle TMB files differently (zip archive)
        if data_type == 'tmb':
            # Define TMB files using guild-specific paths
            tmb_files = [
                (paths['character_file'], 'character-json.json'),
                (paths['attendance_file'], 'hopium-attendance.csv'),
                (paths['item_file'], 'item-notes.csv')
            ]
            
            # Check which files exist
            existing_files = []
            missing_files = []
            total_size = 0
            
            for file_path, filename in tmb_files:
                if os.path.exists(file_path):
                    existing_files.append((file_path, filename))
                    total_size += os.path.getsize(file_path)
                else:
                    missing_files.append(filename)
            
            if not existing_files:
                await ctx.send("❌ No TMB files found. Ensure the TMB directory contains data files.", delete_after=15)
                return
            
            # Create zip file in guild-specific directory
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            zip_filename = f"tmb_data_{ctx.guild.name}_{timestamp}.zip"
            zip_path = os.path.join(paths['sheet_dir'], zip_filename)
            
            try:
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for file_path, filename in existing_files:
                        zipf.write(file_path, filename)
                
                # Create embed for TMB data
                embed = discord.Embed(
                    title="📊 Guild TMB Data",
                    description=f"TMB data files archive containing {len(existing_files)} files",
                    color=0x0066cc,
                    timestamp=datetime.now()
                )
                
                # Add files summary
                file_list = []
                for _, filename in existing_files:
                    file_list.append(f"✅ {filename}")
                
                if missing_files:
                    for filename in missing_files:
                        file_list.append(f"❌ {filename} (missing)")
                
                embed.add_field(
                    name="📋 Files Included",
                    value="\n".join(file_list),
                    inline=False
                )
                
                # Archive info
                zip_size = os.path.getsize(zip_path)
                embed.add_field(
                    name="📁 Archive Info",
                    value=f"**Archive Size:** {zip_size:,} bytes\n**Total Files:** {len(existing_files)}\n**Created:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    inline=True
                )
                
                embed.add_field(
                    name="👤 Requested by",
                    value=ctx.author.mention,
                    inline=True
                )
                
                embed.set_footer(text="This message will be deleted in 60 seconds • Data is sensitive")
                
                # Send the zip file
                await ctx.send(
                    embed=embed,
                    file=discord.File(zip_path, filename=zip_filename),
                    delete_after=60
                )
                
                # Clean up temporary zip file
                os.remove(zip_path)
                
                # Log the action
                print(f"🔒 TMB data downloaded by {ctx.author} ({ctx.author.id}) in {ctx.guild.name}")
                
            except Exception as e:
                # Clean up zip file if it was created
                if os.path.exists(zip_path):
                    os.remove(zip_path)
                raise e
        
        else:
            # Handle single files (armory, icons, parses) using guild-specific paths
            if data_type == 'armory':
                file_path = paths['armory_file']
                file_type = "Armory"
                icon = "🛡️"
            elif data_type == 'icons':
                file_path = paths['item_icons_file']
                file_type = "Item Icons"
                icon = "🖼️"
            elif data_type == 'parses':
                file_path = paths['parses_file']
                file_type = "WCL Parses"
                icon = "📊"
            
            # Check if file exists
            if not os.path.exists(file_path):
                await ctx.send(f"❌ {file_type} file not found. Run the periodic task first to generate data.", delete_after=15)
                return
            
            # Load data
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            if not data:
                await ctx.send(f"ℹ️ {file_type} file is empty. No data available.", delete_after=15)
                return
            
            # Create embed with data summary based on type
            if data_type == 'armory':
                embed = discord.Embed(
                    title=f"{icon} Guild {file_type} Data",
                    description=f"Character equipment data from {len(data)} players",
                    color=0x00ff00,
                    timestamp=datetime.now()
                )
                
                # Add character summary
                total_items = sum(len(items) for items in data.values())
                embed.add_field(
                    name="📊 Summary",
                    value=f"**Players:** {len(data)}\n**Total Items Tracked:** {total_items}",
                    inline=False
                )
            elif data_type == 'icons':
                embed = discord.Embed(
                    title=f"{icon} Guild {file_type} Data",
                    description=f"Item icon data for {len(data)} items",
                    color=0x9932cc,
                    timestamp=datetime.now()
                )
                
                # Add icons summary
                embed.add_field(
                    name="📊 Summary",
                    value=f"**Total Items:** {len(data)}",
                    inline=False
                )
            elif data_type == 'parses':
                embed = discord.Embed(
                    title=f"{icon} Guild {file_type} Data",
                    description=f"WCL performance data for {len(data)} players",
                    color=0xff6600,
                    timestamp=datetime.now()
                )
                
                # Add parses summary
                valid_players = sum(1 for player_data in data.values() if player_data.get("bestPerformanceAverage", 0) > 0)
                avg_best = sum(player_data.get("bestPerformanceAverage", 0) for player_data in data.values()) / len(data) if data else 0
                avg_median = sum(player_data.get("medianPerformanceAverage", 0) for player_data in data.values()) / len(data) if data else 0
                
                embed.add_field(
                    name="📊 Summary",
                    value=f"**Total Players:** {len(data)}\n**Players with Data:** {valid_players}\n**Avg Best Performance:** {avg_best:.1f}\n**Avg Median Performance:** {avg_median:.1f}",
                    inline=False
                )
            
            # File info (common for both types)
            file_size = os.path.getsize(file_path)
            file_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
            embed.add_field(
                name="📁 File Info",
                value=f"**Size:** {file_size:,} bytes\n**Last Updated:** {file_modified.strftime('%Y-%m-%d %H:%M:%S')}",
                inline=True
            )
            
            embed.add_field(
                name="👤 Requested by",
                value=ctx.author.mention,
                inline=True
            )
            
            embed.set_footer(text="This message will be deleted in 60 seconds • Data is sensitive")
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{data_type}_data_{timestamp}.json"
            
            # Send the file as attachment with embed
            message = await ctx.send(
                embed=embed, 
                file=discord.File(file_path, filename=filename),
                delete_after=60
            )
            
            # Log the action
            print(f"🔒 {file_type} data downloaded by {ctx.author} ({ctx.author.id}) in {ctx.guild.name}")
        
        # Delete the command message
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
            
    except json.JSONDecodeError:
        if data_type == 'tmb':
            await ctx.send("❌ Error: One of the TMB files contains invalid JSON data.", delete_after=15)
        else:
            await ctx.send(f"❌ Error: {data_type.title()} file is corrupted or contains invalid JSON.", delete_after=15)
    except FileNotFoundError:
        if data_type == 'tmb':
            await ctx.send("❌ Error: TMB files not found.", delete_after=15)
        else:
            await ctx.send(f"❌ Error: {data_type.title()} file not found.", delete_after=15)
    except Exception as e:
        await ctx.send(f"❌ Error retrieving {data_type} data: {str(e)[:100]}...", delete_after=15)
        print(f"Error in get_file_data command: {e}")

@bot.command(name='get')
async def get_player_data(ctx, data_type: str = None, player_name: str = None):
    """
    Get player armory or parses data as a temporary message
    Usage: !get armory <playerName> | !get parses <playerName>
    Restricted to Officers, Guild Leaders, and authorized roles
    """
    # Check permissions - only allow Officers, Guild Leaders
    authorized_roles = ["Officer", "Guild Leader"]
    user_roles = [role.name for role in ctx.author.roles]
    
    if not any(role in authorized_roles for role in user_roles):
        await ctx.send("❌ You don't have permission to use this command. Required roles: Officer or Guild Leader.", delete_after=10)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    if data_type not in ['armory', 'parses'] or not player_name:
        embed = discord.Embed(
            title="📋 Player Data Lookup",
            description=f"Use `!get <type> <playerName>` to retrieve specific player's data from **{ctx.guild.name}**.",
            color=0xff9900
        )
        embed.add_field(
            name="Valid Commands",
            value="`!get armory <playerName>` - Get specific player's items\n`!get parses <playerName>` - Get specific player's WCL performance data\n`!getfile armory` - Download full guild armory data\n`!getfile parses` - Download full guild parses data\n`!uploadarmory` - Upload armory.json file (merges with existing)",
            inline=False
        )
        embed.add_field(
            name="Examples",
            value="`!get armory Karumenta` - Get Karumenta's items\n`!get parses Karumenta` - Get Karumenta's performance data",
            inline=False
        )
        embed.add_field(
            name="🏰 Server Info",
            value=f"**Guild:** {ctx.guild.name}\n**Data is specific to this server**",
            inline=False
        )
        embed.set_footer(text=f"Guild: {ctx.guild.name} • Data is server-specific")
        await ctx.send(embed=embed, delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Get guild-specific file paths
    guild_id = ctx.guild.id
    paths = get_guild_file_paths(guild_id)
    
    # Initialize guild data if needed
    initialize_guild_data_files(guild_id)
    
    try:
        # Determine which file to use based on data type
        if data_type == 'armory':
            file_path = paths['armory_file']
            file_type = "Armory"
            icon = "🛡️"
        elif data_type == 'parses':
            file_path = paths['parses_file']
            file_type = "WCL Parses"
            icon = "📊"
        
        # Check if file exists
        if not os.path.exists(file_path):
            await ctx.send(f"❌ {file_type} file not found. Run the periodic task first to generate data.", delete_after=15)
            return
        
        # Load data
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        if not data:
            await ctx.send(f"ℹ️ {file_type} file is empty. No character data available.", delete_after=15)
            return
        
        # Find the player (case-insensitive search)
        player_found = None
        player_data = None
        
        for char_name, char_data in data.items():
            if char_name.lower() == player_name.lower():
                player_found = char_name
                player_data = char_data
                break
        
        if not player_found:
            # Search for partial matches
            partial_matches = []
            for char_name in data.keys():
                if player_name.lower() in char_name.lower():
                    partial_matches.append(char_name)
            
            if partial_matches:
                embed = discord.Embed(
                    title="❓ Player Not Found - Did you mean?",
                    description=f"Could not find exact match for '{player_name}'. Did you mean one of these?",
                    color=0xffa500
                )
                
                suggestion_list = []
                for match in partial_matches[:10]:  # Limit to 10 suggestions
                    if data_type == 'armory':
                        data_count = len(data[match]) if isinstance(data[match], list) else 0
                        suggestion_list.append(f"**{match}** ({data_count} items)")
                    elif data_type == 'parses':
                        best_avg = data[match].get("bestPerformanceAverage", 0) if isinstance(data[match], dict) else 0
                        suggestion_list.append(f"**{match}** (Best: {best_avg:.1f})")
                
                embed.add_field(
                    name="Suggestions",
                    value="\n".join(suggestion_list),
                    inline=False
                )
                embed.set_footer(text="Use the exact character name for best results")
            else:
                embed = discord.Embed(
                    title="❌ Player Not Found",
                    description=f"No player found with name '{player_name}' in the {file_type.lower()} data.",
                    color=0xff0000
                )
                
            await ctx.send(embed=embed, delete_after=30)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        # Player found, create embed based on data type
        if data_type == 'armory':
            # Handle armory data (list of items)
            player_items = player_data if isinstance(player_data, list) else []
            
            if not player_items:
                embed = discord.Embed(
                    title=f"📦 {player_found}'s Armory",
                    description="No items found for this player.",
                    color=0xffa500
                )
            else:
                embed = discord.Embed(
                    title=f"🛡️ {player_found}'s Armory",
                    description=f"Found **{len(player_items)}** items for {player_found}",
                    color=0x00ff00,
                    timestamp=datetime.now()
                )
                
                # Format all items as a single list (no chunking)
                item_list = []
                for item in player_items:
                    item_list.append(f"• {item}")
                
                # Join all items into one field value, Discord will handle truncation if needed
                items_text = "\n".join(item_list)
                
                # If the text is too long for a single field, we'll truncate with a note
                if len(items_text) > 1024:
                    # Find a good truncation point (at a line break)
                    truncate_at = items_text.rfind("\n", 0, 1000)
                    if truncate_at == -1:
                        truncate_at = 1000
                    items_text = items_text[:truncate_at] + f"\n... and {len(player_items) - items_text[:truncate_at].count('•')} more items"
                
                embed.add_field(
                    name="🎽 Equipment",
                    value=items_text,
                    inline=False
                )
        
        elif data_type == 'parses':
            # Handle parses data (dictionary with performance metrics)
            if not isinstance(player_data, dict) or not player_data:
                embed = discord.Embed(
                    title=f"📊 {player_found}'s WCL Parses",
                    description="No parse data found for this player.",
                    color=0xffa500
                )
            else:
                # Extract parse data
                best_avg = player_data.get("bestPerformanceAverage", 0)
                median_avg = player_data.get("medianPerformanceAverage", 0)
                
                # Determine color based on performance
                if best_avg >= 95:
                    color = 0xff6600  # Orange (Legendary)
                elif best_avg >= 75:
                    color = 0x9d4edd  # Purple (Epic)
                elif best_avg >= 50:
                    color = 0x0099ff  # Blue (Rare)
                elif best_avg >= 25:
                    color = 0x00ff00  # Green (Uncommon)
                else:
                    color = 0x808080  # Gray (Poor)
                
                embed = discord.Embed(
                    title=f"📊 {player_found}'s WCL Parses",
                    description=f"Performance data for {player_found}",
                    color=color,
                    timestamp=datetime.now()
                )
                
                emoji = "🏆"
                # Performance rating
                if best_avg >= 95:
                    emoji = "🧡"
                elif best_avg >= 75:
                    emoji = "💜"
                elif best_avg >= 50:
                    emoji = "💙 re"
                elif best_avg >= 25:
                    emoji = "💚"
                else:
                    emoji = "🤍"

                # Performance metrics
                embed.add_field(
                    name= emoji+" Performance Averages",
                    value=f"**Best Performance:** {best_avg:.1f}\n**Median Performance:** {median_avg:.1f}",
                    inline=False
                )
                
                # Add any additional parse data if available
                if "encounters" in player_data and isinstance(player_data["encounters"], dict):
                    encounter_list = []
                    for encounter, data in list(player_data["encounters"].items())[:5]:  # Show top 5 encounters
                        if isinstance(data, dict) and "bestPercent" in data:
                            encounter_list.append(f"• **{encounter}**: {data['bestPercent']:.1f}%")
                    
                    if encounter_list:
                        embed.add_field(
                            name="🎯 Top Encounters",
                            value="\n".join(encounter_list),
                            inline=False
                        )
        
        # Add metadata
        file_modified = datetime.fromtimestamp(os.path.getmtime(file_path))
        embed.add_field(
            name="📊 Info",
            value=f"**Last Updated:** {file_modified.strftime('%Y-%m-%d %H:%M:%S')}\n**Requested by:** {ctx.author.mention}",
            inline=False
        )
        
        embed.set_footer(text="This message will be deleted in 45 seconds")
        
        # Send the embed
        await ctx.send(embed=embed, delete_after=45)
        
        # Log the action
        print(f"🔍 Player {file_type.lower()} lookup: {player_found} by {ctx.author} ({ctx.author.id}) in {ctx.guild.name}")
        
        # Delete the command message
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
            
    except json.JSONDecodeError:
        await ctx.send(f"❌ Error: {file_type} file is corrupted or contains invalid JSON.", delete_after=15)
    except FileNotFoundError:
        await ctx.send(f"❌ Error: {file_type} file not found.", delete_after=15)
    except Exception as e:
        await ctx.send(f"❌ Error retrieving player {file_type.lower()} data: {str(e)[:100]}...", delete_after=15)
        print(f"Error in get_player_armory command: {e}")

@bot.command(name='uploadtmb')
async def upload_tmb_files(ctx):
    """
    Upload TMB files (character-json.json, hopium-attendance.csv, item-notes.csv)
    Usage: !uploadtmb (attach 1-3 files)
    Restricted to Officers, Guild Leaders, and authorized roles
    """
    # Check permissions - only allow Officers, Guild Leaders
    authorized_roles = ["Officer", "Guild Leader"]
    user_roles = [role.name for role in ctx.author.roles]
    
    if not any(role in authorized_roles for role in user_roles):
        await ctx.send("❌ You don't have permission to use this command. Required roles: Officer or Guild Leader.", delete_after=10)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Get guild-specific file paths
    guild_id = ctx.guild.id
    paths = get_guild_file_paths(guild_id)
    
    # Initialize guild data if needed
    initialize_guild_data_files(guild_id)
    
    # Check if files are attached
    if not ctx.message.attachments:
        embed = discord.Embed(
            title="📤 TMB File Upload",
            description=f"Upload TMB data files to update the **{ctx.guild.name}** guild database.",
            color=0xff9900
        )
        embed.add_field(
            name="📋 Supported Files",
            value="• `character-json.json` - Character data\n• `hopium-attendance.csv` - Attendance records\n• `item-notes.csv` - Item notes",
            inline=False
        )
        embed.add_field(
            name="📝 Instructions",
            value="1. Attach 1-3 files to your message\n2. Use the `!uploadtmb` command\n3. Files will be validated before overwriting",
            inline=False
        )
        embed.add_field(
            name="⚠️ Important",
            value="Only files with matching names will be updated. Invalid files will be rejected.",
            inline=False
        )
        embed.add_field(
            name="🏰 Server Info",
            value=f"**Guild:** {ctx.guild.name}\n**Files will be uploaded to this server's data only**",
            inline=False
        )
        embed.set_footer(text=f"Guild: {ctx.guild.name} • Data is server-specific")
        await ctx.send(embed=embed, delete_after=30)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Validate file count
    if len(ctx.message.attachments) > 3:
        await ctx.send("❌ Too many files attached. Maximum 3 files allowed (character-json.json, hopium-attendance.csv, item-notes.csv).", delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    try:
        # Define valid TMB files using guild-specific paths
        valid_files = {
            'character-json.json': (paths['character_file'], 'json'),
            'hopium-attendance.csv': (paths['attendance_file'], 'csv'),
            'item-notes.csv': (paths['item_file'], 'csv')
        }
        
        processed_files = []
        validation_errors = []
        uploaded_files = []
        
        # Process each attachment
        for attachment in ctx.message.attachments:
            filename = attachment.filename.lower()
            
            # Check if filename is valid
            if filename not in valid_files:
                validation_errors.append(f"❌ **{attachment.filename}** - Invalid filename. Expected: {', '.join(valid_files.keys())}")
                continue
            
            # Check file size (max 10MB)
            if attachment.size > 10 * 1024 * 1024:
                validation_errors.append(f"❌ **{attachment.filename}** - File too large (max 10MB)")
                continue
            
            target_path, file_type = valid_files[filename]
            
            try:
                # Download file content
                file_content = await attachment.read()
                
                # Validate file content based on type
                if file_type == 'json':
                    try:
                        # Validate JSON structure
                        json_data = json.loads(file_content.decode('utf-8'))
                        
                        # Additional validation for character-json.json
                        if filename == 'character-json.json':
                            if not isinstance(json_data, list):
                                validation_errors.append(f"❌ **{attachment.filename}** - Invalid format: Expected JSON array")
                                continue
                            
                            # Validate each character entry
                            for i, entry in enumerate(json_data):
                                if not isinstance(entry, dict):
                                    validation_errors.append(f"❌ **{attachment.filename}** - Invalid character entry at index {i}")
                                    break
                                if 'name' not in entry:
                                    validation_errors.append(f"❌ **{attachment.filename}** - Missing 'name' field in character entry at index {i}")
                                    break
                            else:
                                # All entries valid
                                processed_files.append((target_path, file_content, attachment.filename))
                        else:
                            # Generic JSON validation passed
                            processed_files.append((target_path, file_content, attachment.filename))
                            
                    except json.JSONDecodeError as e:
                        validation_errors.append(f"❌ **{attachment.filename}** - Invalid JSON format: {str(e)[:100]}")
                        continue
                    except UnicodeDecodeError:
                        validation_errors.append(f"❌ **{attachment.filename}** - Invalid encoding, expected UTF-8")
                        continue
                
                elif file_type == 'csv':
                    try:
                        # Validate CSV structure
                        csv_content = file_content.decode('utf-8')
                        csv_lines = csv_content.strip().split('\n')
                        
                        if not csv_lines or not csv_lines[0].strip():
                            validation_errors.append(f"❌ **{attachment.filename}** - Empty CSV file")
                            continue
                        
                        # Basic CSV validation - check if it can be parsed
                        import io
                        reader = csv.reader(io.StringIO(csv_content))
                        row_count = 0
                        for row in reader:
                            row_count += 1
                            if row_count > 1000:  # Prevent excessive processing
                                break
                        
                        if row_count == 0:
                            validation_errors.append(f"❌ **{attachment.filename}** - No data found in CSV")
                            continue
                        
                        processed_files.append((target_path, file_content, attachment.filename))
                        
                    except UnicodeDecodeError:
                        validation_errors.append(f"❌ **{attachment.filename}** - Invalid encoding, expected UTF-8")
                        continue
                    except Exception as e:
                        validation_errors.append(f"❌ **{attachment.filename}** - CSV validation error: {str(e)[:100]}")
                        continue
                
            except Exception as e:
                validation_errors.append(f"❌ **{attachment.filename}** - Download error: {str(e)[:100]}")
                continue
        
        # Check if any files were processed successfully
        if not processed_files and validation_errors:
            # All files failed validation
            embed = discord.Embed(
                title="❌ Upload Failed",
                description="All files failed validation. No files were updated.",
                color=0xff0000
            )
            embed.add_field(
                name="Validation Errors",
                value="\n".join(validation_errors[:10]),  # Limit to 10 errors
                inline=False
            )
            await ctx.send(embed=embed, delete_after=30)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        # Create backups and update files
        backup_info = []
        updated_files = []
        
        for target_path, file_content, original_filename in processed_files:
            try:
                # Create backup if original file exists
                if os.path.exists(target_path):
                    backup_path = f"{target_path}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                    with open(target_path, 'rb') as original:
                        with open(backup_path, 'wb') as backup:
                            backup.write(original.read())
                    backup_info.append(f"📋 Backup created: {os.path.basename(backup_path)}")
                
                # Write new file
                with open(target_path, 'wb') as f:
                    f.write(file_content)
                
                updated_files.append(original_filename)
                
            except Exception as e:
                validation_errors.append(f"❌ **{original_filename}** - Write error: {str(e)[:100]}")
                continue
        
        # Send success/failure report
        if updated_files:
            embed = discord.Embed(
                title="✅ Upload Successful",
                description=f"Successfully updated {len(updated_files)} TMB file(s)",
                color=0x00ff00,
                timestamp=datetime.now()
            )
            
            embed.add_field(
                name="📁 Updated Files",
                value="\n".join([f"✅ {filename}" for filename in updated_files]),
                inline=False
            )
            
            if backup_info:
                embed.add_field(
                    name="💾 Backups Created",
                    value="\n".join(backup_info),
                    inline=False
                )
            
            if validation_errors:
                embed.add_field(
                    name="⚠️ Validation Errors",
                    value="\n".join(validation_errors[:5]),  # Limit to 5 errors
                    inline=False
                )
            
            embed.add_field(
                name="👤 Uploaded by",
                value=ctx.author.mention,
                inline=True
            )
            
            embed.set_footer(text="Files have been validated and updated successfully")
            
            # Log the action
            print(f"📤 TMB files uploaded by {ctx.author} ({ctx.author.id}) in {ctx.guild.name}: {', '.join(updated_files)}")
        else:
            embed = discord.Embed(
                title="❌ Upload Failed",
                description="No files were updated due to validation errors.",
                color=0xff0000
            )
            embed.add_field(
                name="Validation Errors",
                value="\n".join(validation_errors[:10]),
                inline=False
            )
        
        await ctx.send(embed=embed, delete_after=60)
        
        # Delete the command message
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
            
    except Exception as e:
        await ctx.send(f"❌ Error processing TMB file uploads: {str(e)[:100]}...", delete_after=15)
        print(f"Error in upload_tmb_files command: {e}")

@bot.command(name='uploadarmory')
async def upload_armory_file(ctx):
    """
    Upload armory.json file to merge with existing armory data
    Usage: !uploadarmory (attach armory.json file)
    Restricted to Officers, Guild Leaders, and authorized roles
    """
    # Check permissions - only allow Officers, Guild Leaders
    authorized_roles = ["Officer", "Guild Leader"]
    user_roles = [role.name for role in ctx.author.roles]
    
    if not any(role in authorized_roles for role in user_roles):
        await ctx.send("❌ You don't have permission to use this command. Required roles: Officer or Guild Leader.", delete_after=10)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Get guild-specific file paths
    guild_id = ctx.guild.id
    paths = get_guild_file_paths(guild_id)
    
    # Initialize guild data if needed
    initialize_guild_data_files(guild_id)
    
    # Check if file is attached
    if not ctx.message.attachments:
        embed = discord.Embed(
            title="📤 Armory File Upload",
            description=f"Upload an armory.json file to merge with **{ctx.guild.name}** guild armory data.",
            color=0xff9900
        )
        embed.add_field(
            name="📋 File Requirements",
            value="• File must be named `armory.json`\n• Must contain valid JSON format\n• Data structure: `{\"PlayerName\": [\"Item1\", \"Item2\"]}`",
            inline=False
        )
        embed.add_field(
            name="📝 Instructions",
            value="1. Attach the `armory.json` file to your message\n2. Use the `!uploadarmory` command\n3. File will be validated and merged with existing data",
            inline=False
        )
        embed.add_field(
            name="⚠️ Merge Behavior",
            value="• New players will be added\n• New items for existing players will be added\n• Duplicate items will be ignored\n• Existing data will be preserved",
            inline=False
        )
        embed.add_field(
            name="💾 Backup",
            value="A timestamped backup of the current armory file will be created before merging.",
            inline=False
        )
        embed.add_field(
            name="🏰 Server Info",
            value=f"**Guild:** {ctx.guild.name}\n**File will be merged with this server's data only**",
            inline=False
        )
        embed.set_footer(text=f"Guild: {ctx.guild.name} • Data is server-specific")
        await ctx.send(embed=embed, delete_after=30)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Validate only one file
    if len(ctx.message.attachments) > 1:
        await ctx.send("❌ Please attach only one armory.json file.", delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    attachment = ctx.message.attachments[0]
    
    # Validate filename
    if attachment.filename.lower() != 'armory.json':
        await ctx.send("❌ File must be named `armory.json`. Please rename your file and try again.", delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    # Check file size (max 50MB for armory files)
    if attachment.size > 50 * 1024 * 1024:
        await ctx.send("❌ File too large (max 50MB). Please check your armory file.", delete_after=15)
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass
        return
    
    try:
        # Download and validate file content
        file_content = await attachment.read()
        
        try:
            # Parse JSON content
            uploaded_armory = json.loads(file_content.decode('utf-8'))
        except json.JSONDecodeError as e:
            embed = discord.Embed(
                title="❌ Invalid JSON Format",
                description="The uploaded file contains invalid JSON.",
                color=0xff0000
            )
            embed.add_field(
                name="Error Details",
                value=f"```{str(e)[:500]}```",
                inline=False
            )
            await ctx.send(embed=embed, delete_after=30)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        except UnicodeDecodeError:
            await ctx.send("❌ File encoding error. Please ensure the file is saved as UTF-8.", delete_after=15)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        # Validate armory data structure
        if not isinstance(uploaded_armory, dict):
            await ctx.send("❌ Invalid armory format. Expected JSON object with player names as keys.", delete_after=15)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        validation_errors = []
        valid_players = 0
        total_items = 0
        
        # Validate each player entry
        for player_name, items in uploaded_armory.items():
            if not isinstance(player_name, str) or not player_name.strip():
                validation_errors.append(f"Invalid player name: {repr(player_name)}")
                continue
            
            if not isinstance(items, list):
                validation_errors.append(f"Player '{player_name}': Items must be a list, got {type(items).__name__}")
                continue
            
            # Validate items
            for i, item in enumerate(items):
                if not isinstance(item, str):
                    validation_errors.append(f"Player '{player_name}', item {i+1}: Expected string, got {type(item).__name__}")
                    break
                if not item.strip():
                    validation_errors.append(f"Player '{player_name}', item {i+1}: Empty item name")
                    break
            else:
                # All items valid for this player
                valid_players += 1
                total_items += len(items)
        
        # Check if we have critical validation errors
        if validation_errors and valid_players == 0:
            embed = discord.Embed(
                title="❌ Validation Failed",
                description="The armory file contains critical errors and cannot be processed.",
                color=0xff0000
            )
            embed.add_field(
                name="Errors Found",
                value="\n".join(validation_errors[:10]) + ("..." if len(validation_errors) > 10 else ""),
                inline=False
            )
            await ctx.send(embed=embed, delete_after=30)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        # Load existing armory data using guild-specific paths
        existing_armory = {}
        if os.path.exists(paths['armory_file']):
            try:
                with open(paths['armory_file'], 'r', encoding='utf-8') as f:
                    existing_armory = json.load(f)
            except (json.JSONDecodeError, FileNotFoundError):
                existing_armory = {}
        
        # Create backup before merging using guild-specific paths
        backup_created = False
        backup_path = None
        if os.path.exists(paths['armory_file']):
            try:
                backup_path = f"{paths['armory_file']}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                with open(paths['armory_file'], 'rb') as original:
                    with open(backup_path, 'wb') as backup:
                        backup.write(original.read())
                backup_created = True
            except Exception as e:
                await ctx.send(f"❌ Failed to create backup: {str(e)[:100]}...", delete_after=15)
                try:
                    await ctx.message.delete()
                except discord.NotFound:
                    pass
                return
        
        # Merge armory data
        merge_stats = {
            'new_players': 0,
            'updated_players': 0,
            'new_items': 0,
            'duplicate_items': 0,
            'total_players_processed': 0
        }
        
        merged_armory = existing_armory.copy()
        
        for player_name, items in uploaded_armory.items():
            # Skip invalid entries
            if not isinstance(player_name, str) or not isinstance(items, list):
                continue
            
            clean_player_name = player_name.strip()
            if not clean_player_name:
                continue
            
            # Clean and validate items
            clean_items = []
            for item in items:
                if isinstance(item, str) and item.strip():
                    clean_items.append(item.strip())
            
            if not clean_items:
                continue
            
            merge_stats['total_players_processed'] += 1
            
            if clean_player_name not in merged_armory:
                # New player
                merged_armory[clean_player_name] = clean_items
                merge_stats['new_players'] += 1
                merge_stats['new_items'] += len(clean_items)
            else:
                # Existing player - merge items
                existing_items = set(merged_armory[clean_player_name])
                new_items = []
                
                for item in clean_items:
                    if item in existing_items:
                        merge_stats['duplicate_items'] += 1
                    else:
                        new_items.append(item)
                        merge_stats['new_items'] += 1
                
                if new_items:
                    merged_armory[clean_player_name].extend(new_items)
                    merge_stats['updated_players'] += 1
        
        # Save merged armory data using guild-specific paths
        try:
            # Sort merged armory data alphabetically by character name before saving
            sorted_merged_armory = dict(sorted(merged_armory.items()))
            
            # Write to temporary file first for atomic operation
            temp_file = paths['armory_file'] + '.tmp'
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(sorted_merged_armory, f, ensure_ascii=False, indent=2)
            
            # Atomic rename
            os.replace(temp_file, paths['armory_file'])
            
        except Exception as e:
            # Clean up temp file if it exists
            if os.path.exists(temp_file):
                os.remove(temp_file)
            await ctx.send(f"❌ Failed to save merged armory data: {str(e)[:100]}...", delete_after=15)
            try:
                await ctx.message.delete()
            except discord.NotFound:
                pass
            return
        
        # Create success report
        embed = discord.Embed(
            title="✅ Armory Upload & Merge Successful",
            description="Armory data has been successfully merged with existing data.",
            color=0x00ff00,
            timestamp=datetime.now()
        )
        
        # Merge statistics
        stats_text = []
        if merge_stats['new_players'] > 0:
            stats_text.append(f"👤 **{merge_stats['new_players']}** new players added")
        if merge_stats['updated_players'] > 0:
            stats_text.append(f"📝 **{merge_stats['updated_players']}** existing players updated")
        if merge_stats['new_items'] > 0:
            stats_text.append(f"⚔️ **{merge_stats['new_items']}** new items added")
        if merge_stats['duplicate_items'] > 0:
            stats_text.append(f"🔄 **{merge_stats['duplicate_items']}** duplicate items skipped")
        
        if stats_text:
            embed.add_field(
                name="📊 Merge Statistics",
                value="\n".join(stats_text),
                inline=False
            )
        
        # File info
        embed.add_field(
            name="📁 File Information",
            value=f"**Size:** {attachment.size:,} bytes\n**Players Processed:** {merge_stats['total_players_processed']}\n**Total Items in Upload:** {total_items}",
            inline=True
        )
        
        # Backup info
        if backup_created:
            embed.add_field(
                name="💾 Backup Created",
                value=f"`{os.path.basename(backup_path)}`",
                inline=True
            )
        
        # Validation warnings
        if validation_errors:
            embed.add_field(
                name="⚠️ Validation Warnings",
                value=f"{len(validation_errors)} entries skipped due to validation errors.\nProcessed {valid_players} valid players.",
                inline=False
            )
        
        embed.add_field(
            name="👤 Uploaded by",
            value=ctx.author.mention,
            inline=True
        )
        
        embed.set_footer(text="Armory data merged successfully • Use !get armory <player> to view player items")
        
        await ctx.send(embed=embed, delete_after=60)
        
        # Log the action
        print(f"📤 Armory file uploaded and merged by {ctx.author} ({ctx.author.id}) in {ctx.guild.name}")
        print(f"   Stats: {merge_stats['new_players']} new players, {merge_stats['new_items']} new items, {merge_stats['duplicate_items']} duplicates skipped")
        
    except Exception as e:
        await ctx.send(f"❌ Error processing armory upload: {str(e)[:100]}...", delete_after=15)
        print(f"Error in upload_armory_file command: {e}")
    finally:
        # Always delete the command message
        try:
            await ctx.message.delete()
        except discord.NotFound:
            pass

def createExcel(guild_id, excelType):
    """Create Excel file with guild-specific data"""
    # Get guild-specific file paths
    paths = get_guild_file_paths(guild_id)
    
    # Initialize guild data if needed
    initialize_guild_data_files(guild_id)
    players = {}
    wowClasses = []

    iconWidth = 4
    iconHeight = 25

    itemsIcons = {}
    playerParse = {}

    # Get Blizzard API token with retry logic
    access_token = None
    for attempt in range(3):
        try:
            response = requests.post(
                BLIZZARD_TOKEN_URL, 
                data={'grant_type': 'client_credentials'}, 
                auth=(BLIZZARD_ID, BLIZZARD_SECRET),
                timeout=10
            )
            response.raise_for_status()
            access_token = response.json()['access_token']
            break
        except requests.RequestException as e:
            print(f"⚠️ Token request attempt {attempt + 1} failed: {e}")
            if attempt == 2:
                print("❌ Failed to get Blizzard API token after 3 attempts")
                return

    with open(paths['item_icons_file'], 'r', encoding='utf-8') as f:
        try:
            itemsIcons = json.load(f)
            if not itemsIcons:
                itemsIcons = {}
        except json.JSONDecodeError:
            print("Warning: item-icons.json is corrupted or invalid. Starting with empty icons map.")
            itemsIcons = {}

    with open(paths['parses_file'], 'r', encoding='utf-8') as f:
        try:
            playerParse = json.load(f)
            if not playerParse:
                playerParse = {}
        except json.JSONDecodeError:
            print("Warning: parses.json is corrupted or invalid. Starting with empty parses data.")
            playerParse = {}

    #Attendance Start
    with open(paths['attendance_file'], newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='|')
        
        firstRow = next(csvreader)
        attendanceDates = []
        
        for row in csvreader:
            date = datetime.strptime(row[0].replace('"', ''), "%Y-%m-%d %H:%M:%S").strftime("%d/%m/%y")
            if not date in attendanceDates:
                attendanceDates.append(date)
            
            playerName = row[2].replace('"', '').capitalize()
            player = {}
            raids = []
            benchedRaids = []
            absentRaids = []
            unpreparedRaids = []

            try: 
                if players[playerName]:
                    player = players[playerName]
                    raids = player["raids"]
                    benchedRaids = player["benched_raids"]
                    absentRaids = player["absent_raids"]
                    unpreparedRaids = player["unprepared_raids"]
            except:
                player["name"] = playerName
            
            if row[6].replace('"', '') == "Benched":
                benchedRaids.append(date)
            elif row[6].replace('"', '') == "Gave notice":
                absentRaids.append(date)
            elif row[6].replace('"', '') == "Unprepared":
                unpreparedRaids.append(date)
            else:
                raids.append(date)
            
            player["raids"] = raids
            player["benched_raids"] = benchedRaids
            player["absent_raids"] = absentRaids
            player["unprepared_raids"] = unpreparedRaids
            player["firstRaid"] = date
            player["isInAttendance"] = True
            players[playerName] = player
    #Attendance Finish
    
    #Loot Start
    playerData = ""
    with open(paths['character_file'], 'r', encoding='utf-8') as file:
        playerData = json.load(file)
    
    for playerInfo in playerData :
        player = {}
        name = ""
        try:
            name = playerInfo["name"].capitalize()
            player = players[name]
        except:
            player["name"] = playerInfo["name"].capitalize()
            player["role"] = playerInfo["archetype"]
            player["firstRaid"] = "31/12/30"
            player["raids"] = []
            player["benched_raids"] = []
            player["absent_raids"] = []
            player["unprepared_raids"] = []
            player["wishlist"] = "0/0"
            print("Player " + name + " is not present in the attendance file. Probably a new player or it's unclaimed on TMB.")
        
        player["isInLootData"] = True
        loot = {}
        
        for lootReceived in playerInfo["received"]:
            lootItem = {}
            lootItem["name"] = lootReceived["name"]
            lootItem["id"] = lootReceived["item_id"]
            lootItem["isOS"] = lootReceived["pivot"]["is_offspec"]
            
            date_object = datetime.strptime(lootReceived["pivot"]["received_at"], "%Y-%m-%d %H:%M:%S")
            formatted_date = date_object.strftime("%d/%m/%y")
            lootItem["receivedDate"] = formatted_date
            loot[lootItem["name"]] = lootItem
            
        wishlist = []
        sum = 0
        try:
            wishlist = playerInfo["wishlist"]
        except:
            print("No wishlist found for " + name)
        for wishlistItem in playerInfo["wishlist"]:
            if wishlistItem["pivot"]["is_received"] == 1:
                sum += 1
        
        if playerInfo["class"] is not None and not playerInfo["class"] in wowClasses:
            wowClasses.append(playerInfo["class"])

        playerParseData = playerParse.get(name, {})
        player["bestPerformanceAverage"] = playerParseData.get("bestPerformanceAverage", 0.0)
        player["medianPerformanceAverage"] = playerParseData.get("medianPerformanceAverage", 0.0)

        player["race"] = playerInfo["race"]
        player["role"] = playerInfo["archetype"]
        player["class"] = playerInfo["class"]
        player["is_alt"] = playerInfo["is_alt"]
        player["member_id"] = playerInfo["member_id"]
        player["wishlist"] = str(sum)+"/"+str(len(wishlist))
        player["loot"] = loot
        
        players[name] = player
    #Loot Finish
    
    #First Sheet Start
    column_names = ["Name", "Class", "Race", "Raids", "Benched", "Attendance", "Items (+OS)", "MS Ratio", "Last MS", "Wishlist", "Best avg parse", "Median avr parse", "Last bench", "Name"]
    counter = 0
    for date in attendanceDates :
        counter += 1
        if counter == 20:
            counter = 0
            column_names.append("Name")
        column_names.append(date)   
    
    playerInfoList = []
    
    for player in list(players.values()):
        try:
            if not player["isInLootData"]:
                print("Removed player " + player["name"] + " since he's not in the roster.")
                del players[player["name"]]
                continue
        except:
            print("Removed player " + player["name"] + " since he's not in the roster.")
            del players[player["name"]]
            continue
            
        try:
            if player["is_alt"]:
                print("Removed player " + player["name"] + " since he's an alt.")
                del players[player["name"]]
                continue
        except:
            print("Removed player " + player["name"] + " since he's an alt.")
            del players[player["name"]]
            continue
            
        playerInfo = []
        playerDateInfo = []
        totalRaids = 0
        completedRaids = 0
        benchedRaids = 0
        absentRaids = 0
        attendance = 0.0
        itemsPlaceholders = "<MS>(+<OS>)"
        lastReceivedItemDate = "-"
        lastBench = "-"
        msItems = 0
        msRatio = 0.0
        osItems = 0
        
        counter = 0
        for date in attendanceDates :

            counter += 1
            if counter == 20:
                counter = 0
                try:
                    name = player["name"].capitalize()
                    if player["member_id"] is None :
                        name += " (Unclaimed)"
                    playerDateInfo.append(name)
                except:
                    playerDateInfo.append("-")

            found = False
            benched = False
            unprepared_b = False
            
            for raid in player["raids"]:
                if date == raid :
                    completedRaids += 1
                    found = True
                
            for bench in player["benched_raids"]:
                if date == bench :
                    benchedRaids += 1
                    benched = True
                    
            for unprepared in player["unprepared_raids"]:
                if date == unprepared :
                    unprepared_b = True
                
            if not found:
                if benched:
                    if lastBench == "-":
                        lastBench = date
                    playerDateInfo.append("Benched")
                elif unprepared_b:
                    playerDateInfo.append("Holiday")
                else:
                    if datetime.strptime(player["firstRaid"], "%d/%m/%y") > datetime.strptime(date, "%d/%m/%y"):
                        playerDateInfo.append("N/A")
                    else:
                        absentRaids += 1
                        playerDateInfo.append("Absent")
            else:
                currentMsItems = 0
                currentOsItems = 0
                
                lootReceived = {}
                try:
                    lootReceived = player["loot"]
                except:
                    lootReceived = {}
                    
                for loot in lootReceived.values():
                    if loot["receivedDate"] == date:
                        if loot["isOS"] == 0:
                            if lastReceivedItemDate == "-":
                                lastReceivedItemDate = date
                            currentMsItems += 1
                        else:
                            currentOsItems += 1
                if currentMsItems == 0 and currentOsItems == 0:
                    appendStr = "-"
                else:
                    appendStr = itemsPlaceholders.replace("<MS>", str(currentMsItems)).replace("<OS>", str(currentOsItems))
                playerDateInfo.append(appendStr)
                
                msItems += currentMsItems
                osItems += currentOsItems
        
        if completedRaids > 0:
            msRatio = f"{msItems / completedRaids:.2f}"
            osRatio = f"{osItems / completedRaids:.2f}"
        
        totalRaids = completedRaids + benchedRaids + absentRaids

        #Name
        try:
            name = player["name"].capitalize()
            if player["member_id"] is None :
                name += " (Unclaimed)"
            playerInfo.append(name)
        except:
            playerInfo.append("-")
        #Class
        try:
            playerInfo.append(player["class"])
        except:
            playerInfo.append("-")
        #Race
        try:
            playerInfo.append(player["race"])
        except:
            playerInfo.append("-")
        #Completed Raids
        playerInfo.append(completedRaids)
        #Benched Raids
        playerInfo.append(benchedRaids)
        #Attendance
        if totalRaids > 0:
            attendance = ((completedRaids + benchedRaids) / totalRaids)
        player["attendance"] = f"{attendance*100:.1f}" 
        playerInfo.append(attendance)
        #MS Items
        playerInfo.append(str(msItems) + " (" + str(osItems) + ")")
        playerInfo.append(msRatio)
        player["msRatio"] = msRatio
        #Last Received Item
        playerInfo.append(lastReceivedItemDate)
        #Wishlist
        try:
            playerInfo.append(player["wishlist"])
        except:
            playerInfo.append("0/0")
        #Parse Info
        best_avg = player.get("bestPerformanceAverage", 0)
        median_avg = player.get("medianPerformanceAverage", 0)
        
        # Ensure values are numeric, default to 0 if not
        try:
            best_avg = float(best_avg) if best_avg not in [None, "", "-"] else 0
        except (ValueError, TypeError):
            best_avg = 0
            
        try:
            median_avg = float(median_avg) if median_avg not in [None, "", "-"] else 0
        except (ValueError, TypeError):
            median_avg = 0
        
        playerInfo.append(f"{best_avg:.1f}" if best_avg != 0 else "-")
        playerInfo.append(f"{median_avg:.1f}" if median_avg != 0 else "-")
        #Last bench
        playerInfo.append(lastBench)
        #Spacer
        try:
            name = player["name"].capitalize()
            if player["member_id"] is None :
                name += " (Unclaimed)"
            playerInfo.append(name)
        except:
            playerInfo.append("-")
        
        #Raids Presence and Loot
        for dateInfo in playerDateInfo:
            playerInfo.append(dateInfo)
        
        #Add
        playerInfoList.append(playerInfo)
    
    workbook = Workbook()
    sheet = workbook.active
    
    if excelType == "Attendance" or excelType == "All" :
        sheet.title = "Attendance"
        sheet.auto_filter.ref = "A1:BZ1"
    
        sorted_data = sorted(playerInfoList, key=lambda x: (x[1], x[5]), reverse=True)
            
        for col_num, column_name in enumerate(column_names, start=1):
            cell = sheet.cell(row=1, column=col_num, value=column_name)
            cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
            cell.font = Font(name="Aptos", bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            #cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

        for row_num, row_data in enumerate(sorted_data, start=2):
            for col_num, cell_value in enumerate(row_data, start=1):
                cell = sheet.cell(row=row_num, column=col_num, value=cell_value)
                
                cell.alignment = Alignment(horizontal="center")        
                if col_num == 4 or col_num == 5 or col_num == 8 or col_num == 11 or col_num == 12:
                    for row in range(2, len(players) + 2):
                        sheet.cell(row=row, column=col_num).number_format = "0"  # Numeric format
                if col_num == 6:
                    for row in range(2, len(players) + 2):
                        sheet.cell(row=row, column=col_num).number_format = "0.00%"
                if col_num == 9 or col_num == 13:
                    for row in range(2, len(players) + 2):
                        sheet.cell(row=row, column=col_num).number_format = "DD/MM"


        column_sizes = {
            "Name": 22,
            "Class": 12,
            "Race": 12,
            "Raids": 14,
            "Benched": 14,
            "Attendance": 16,
            "Items (+OS)": 14,
            "MS Ratio": 14,
            "Last MS": 14,
            "Whishlist": 14,
            "Best avg parse": 18,
            "Median avr parse": 18,
            "OS Items": 16,
            "OS Ratio": 14,
            "Last Bench": 16,
            "Default": 14
        }
        # Adjust column widths
        for col_num, column_name in enumerate(column_names, start=1):
            column_letter = get_column_letter(col_num)
            
            try:
                size = column_sizes[column_name]
            except:
                size = column_sizes["Default"]
                        
            sheet.column_dimensions[column_letter].width = max(len(column_name), size)
            
            for row in range(2, 80):
                cell = sheet[column_letter + str(row)]
                if cell.value is not None:
                    cell.alignment = Alignment(horizontal="center")

                    thin_border = Side(border_style="thin", color="000000")  # Black thin border
                    cell.border = Border(top=thin_border, bottom=thin_border, left=thin_border, right=thin_border)
                    
                    cell.font = Font(name="Aptos Light", bold=False)
                        
                    if column_name == "Name" or column_name == "Class":
                        value = cell.value
                        if column_name == "Name":
                            cell.alignment = Alignment(horizontal="left")
                            cell.font = Font(name="Aptos", bold=True)
                            value = sheet["B" + str(row)].value
                        bgcolor = CLASS_LIST.get(value, {}).get("color", "CCCCCC")
                        cell.fill = PatternFill(start_color=bgcolor, end_color=bgcolor, fill_type="solid")
                    elif column_name == "Race":
                        bgcolor = "CCCCCC"
                        if cell.value == "Dwarf":
                            bgcolor = "C69B6D"
                        elif cell.value == "Gnome":
                            bgcolor = "FFF468"
                        elif cell.value == "Human":
                            bgcolor = "F48CBA"
                        elif cell.value == "Night Elf":
                            bgcolor = "0070DD"
                        elif cell.value == "Orc":
                            bgcolor = "AAD372"
                        elif cell.value == "Tauren":
                            bgcolor = "C69B6D"
                        elif cell.value == "Troll":
                            bgcolor = "FFFFFF"
                        elif cell.value == "Undead":
                            bgcolor = "8788EE"
                        cell.fill = PatternFill(start_color=bgcolor, end_color=bgcolor, fill_type="solid")
                    elif column_name == "Attendance":
                        start_color = (255, 156, 0)
                        end_color = (117, 249, 77)
                        percentage = cell.value
                        color = calculate_gradient_color(percentage, start_color, end_color)
                        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    elif column_name == "MS Ratio":
                        start_color = (255, 255, 255)
                        end_color = (186, 72, 177) 
                        if cell.value != "-" and cell.value is not None:
                            percentage = float(cell.value)
                            color = calculate_gradient_color(percentage, start_color, end_color)
                            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                    elif column_name == "Wishlist":
                        if cell.value == "0/0":
                            cell.value = "Empty"
                            cell.fill = PatternFill(start_color="FF9C00", end_color="FF9C00", fill_type="solid")
                        elif cell.value.split("/")[0] == cell.value.split("/")[1]:
                            cell.fill = PatternFill(start_color="75F94D", end_color="75F94D", fill_type="solid")
                    elif column_name == "Best avg parse" or column_name == "Median avr parse":
                        if cell.value != "-" and cell.value is not None:
                            start_color = (255, 255, 255)
                            end_color = (66, 133, 244) 
                            value = float(cell.value)
                            color = "66664E"
                            if value == 100:
                                color = "E5A93F"
                            elif value >= 99:
                                color = "BE49A8"
                            elif value >= 95:
                                color = "FF8000"
                            elif value >= 75:
                                color = "A335EE"
                            elif value >= 50:
                                color = "0961FE"
                            elif value >= 25:
                                color = "0961FE"
                            cell.font = Font(name="Aptos", bold=True, color=color)
                    elif col_num > 14: # N
                        bgcolor = "CCCCCC"
                        cell.font = Font(name="Aptos Light", bold=True)
                        if cell.value == "N/A":
                            bgcolor = "FFFFFF"
                            cell.value = ""
                        elif cell.value == "Benched":
                            bgcolor = "9DC0FA"
                        elif cell.value == "Absent":
                            bgcolor = "FF9C00"
                        elif cell.value == "Holiday":
                            bgcolor = "00FFFF"
                        elif cell.value == "-":
                            bgcolor = "A1FB8E"
                        else:
                            bgcolor = "75F94D"
                            if cell.value.startswith("0"):
                                cell.font = Font(name="Aptos Light", bold=False)
                        cell.fill = PatternFill(start_color=bgcolor, end_color=bgcolor, fill_type="solid")
        print("Create attendance sheet with " + str(len(players)) + " players.")
    else:
        # If we're not creating an attendance sheet, remove the empty active sheet
        # We'll add other sheets later, so this prevents an empty first sheet
        workbook.remove(sheet)
    #First Sheet Stop

    # Item Sheets Start
    # Loading file
    itemList = {}

    with open(paths['item_file'], newline='', encoding='utf-8') as csvfile:
        csvreader = csv.reader(csvfile, delimiter=',', quotechar='"')
        
        firstRow = next(csvreader)
        
        for row in csvreader:
            itemName = row[0].replace('"', '')
            itemId = row[1]
            itemInstance = row[2].replace('"', '')
            itemSource = row[3].replace('"', '')

            itemNotes = row[5].replace('"', '')
            itemOffNotes = row[6].replace('"', '')
            itemTier = row[7].replace('"', '')

            item = {}
            item["itemName"] = itemName
            item["itemId"] = itemId
            item["itemInstance"] = itemInstance
            item["itemSource"] = itemSource
            item["itemNotes"] = itemNotes
            item["itemOffNotes"] = itemOffNotes
            item["itemTier"] = itemTier

            if itemTier and itemOffNotes:
                itemList[itemId] = item
    # Loading file finish

    #Load armory cache
    with open(paths['armory_file'], 'r', encoding='utf-8') as f:
        try:
            armoryList = json.load(f)
            if not armoryList:
                armoryList = {}
        except json.JSONDecodeError:
            print("Warning: armory.json is corrupted or invalid. Starting with empty armory list.")
            armoryList = {}


    if excelType == "Loot" or excelType == "All" :
        # Create loot items sheet
        allLootSheet = workbook.create_sheet(title="Loot")
        raidList = {
            "Molten Core": "E26B0A",
            "Blackwing Lair": "C0504D",
            "Temple of Ahn'Qiraj": "4F6228",
            "Naxxramas": "403151"
            }
        
        # Create role to classes mapping
        roleToClasses = {}
        for class_name, class_info in CLASS_LIST.items():
            for role in class_info["roles"]:
                if role not in roleToClasses:
                    roleToClasses[role] = []
                roleToClasses[role].append(class_name)

        i = 1
        for raid in raidList.keys():
            raidItems = {}
            for itemId, item in itemList.items():
                if item["itemInstance"] == raid:
                    raidItems[itemId] = item
            if not raidItems:
                continue
            allLootSheet.merge_cells(start_row=i, start_column=1, end_row=i, end_column=5)
            cell = allLootSheet.cell(row=i, column=1, value=raid)
            cell.fill = PatternFill(start_color=raidList[raid], end_color=raidList[raid], fill_type="solid")
            i += 1
            for raidItem in raidItems.values():
                raidItemName = raidItem["itemName"]
                itemOffNotes = raidItem["itemOffNotes"]
                lootRow = ["", raidItem["itemId"], raidItemName, raidItem["itemNotes"] ,raidItem["itemTier"], "","",""]
                for player in players.values():
                    # Parse itemOffNotes to extract classes and roles
                    raidItemClasses = []
                    raidItemRoles = []

                    # Extract classes from < >
                    class_matches = re.findall(r'<([^>]+)>', itemOffNotes)
                    for match in class_matches:
                        raidItemClasses.append(match.strip())

                    # Extract roles from [ ]
                    role_matches = re.findall(r'\[([^\]]+)\]', itemOffNotes)
                    for match in role_matches:
                        raidItemRoles.append(match.strip())

                    # Get player info
                    player_class = player.get("class")
                    player_class_info = CLASS_LIST.get(player_class, {})
                    player_role = player.get("role")

                    # Simplified filtering logic
                    can_use_item = False

                    if raidItemClasses and not raidItemRoles:
                        # Only classes specified - filter by class
                        can_use_item = player_class in raidItemClasses

                    elif raidItemRoles and not raidItemClasses:
                        # Only roles specified - filter by role
                        can_use_item = player_role in raidItemRoles

                    elif raidItemClasses and raidItemRoles:
                        # Both classes and roles specified - filter by roles + class compatibility
                        class_match = player_class in raidItemClasses
                        role_match = player_role in raidItemRoles
                        class_roles = player_class_info.get("roles", [])

                        isRoleToFilter = any(role in class_roles for role in raidItemRoles)

                        if class_match and role_match:
                            can_use_item = True

                        elif class_match and not isRoleToFilter:
                            can_use_item = True 

                        elif not class_match:
                            for otherClass in raidItemClasses:
                                player_class_info = CLASS_LIST.get(otherClass, {})
                                other_class_roles = player_class_info.get("roles", [])
                                can_use_item = True
                                if player_role in other_class_roles:
                                    can_use_item = False
                                    break

                    if not can_use_item:
                        continue

                    playerName = player["name"].capitalize()
                    found = False
                    try:
                        playerArmory = armoryList[playerName]
                    except KeyError:
                        playerArmory = []
                        armoryList[playerName] = playerArmory
                    for armoryItem in armoryList[playerName]:
                        if armoryItem == raidItemName:
                            found = True
                        elif raidItemName == "Head of Nefarian":
                            if armoryItem == "Master Dragonslayer's Medallion" or armoryItem == "Master Dragonslayer's Orb" or armoryItem == "Master Dragonslayer's Ring":
                                found = True
                        elif raidItemName == "Eye of C'Thun":
                            if armoryItem == "Amulet of the Fallen God" or armoryItem == "Cloak of the Fallen God" or armoryItem == "Ring of the Fallen God":
                                found = True
                        elif raidItemName == "Vek'nilash's Circlet":
                            if armoryItem == "Conqueror's Crown" or armoryItem == "Doomcaller's Circlet" or armoryItem == "Enigma Circlet" or armoryItem == "Tiara of the Oracle":
                                found = True

                    for loot in player["loot"].values():
                        if loot["name"] == raidItemName:
                            found = True
                    if not found:
                        lootRow.append(f'{playerName} ({player["attendance"]}% - {player["msRatio"]})')
                    
                allLootSheet.append(lootRow)
                i += 1

        instanceColor = raidList.get("Molten Core", "E26B0A")
        thin = Side(border_style="thin", color="000000")
        for row in allLootSheet.iter_rows(min_row=1, max_row=allLootSheet.max_row):
            if row[0].value is None or row[0].value == "" and row[1].value is None or row[1].value == "":
                allLootSheet.row_dimensions[row[0].row].height = iconHeight
                continue

            if row[0].value is not None and row[0].value != "":
                allLootSheet.row_dimensions[row[0].row].height = iconHeight
                row[0].font = Font(name="Aptos", bold=True, color="FFFFFF",  size=16)
                row[0].alignment = Alignment(horizontal="center", vertical="center")
                instanceColor = row[0].fill.start_color.index
                continue

            row[0].fill = PatternFill(start_color=instanceColor, end_color=instanceColor, fill_type="solid")
            item_id_cell = row[1]  # Assuming itemId is in the first column
            item_id = str(item_id_cell.value)
            allLootSheet.row_dimensions[item_id_cell.row].height = iconHeight
                
            if item_id not in itemsIcons.keys():
                try:
                    media_url = f'https://eu.api.blizzard.com/data/wow/media/item/{item_id}?namespace=static-classic-eu&locale=en_GB'
                    urlHeaders = {'Authorization': f'Bearer {access_token}'}
                    media_response = requests.get(media_url, headers=urlHeaders)
                    icon_url = media_response.json()['assets'][0]['value']
                    itemsIcons[item_id] = icon_url
                except:
                    print("Error fetching media for loot item:", loot["name"])

            icon_url = itemsIcons.get(item_id)
            if icon_url:
            # Set the cell to use the IMAGE formula
                item_id_cell.value = f'=IMAGE("{icon_url}", 2)'
            row[1].border = Border(left=thin, right=thin, top=thin, bottom=thin)

            row[2].alignment = Alignment(horizontal="left", vertical="center")
            row[2].font = Font(name="Aptos", bold=True)
            row[2].border = Border(left=thin, right=thin, top=thin, bottom=thin)

            if row[3].value is not None and row[3].value != "":
                notes = row[3].value
                row[3].value = '=IMAGE("https://render.worldofwarcraft.com/classic-eu/icons/56/inv_misc_questionmark.jpg", 2)'
                row[3].comment = Comment(text=notes, author="")
            else:
                allLootSheet.merge_cells(start_row=item_id_cell.row, start_column=3, end_row=item_id_cell.row, end_column=4)
            row[3].border = Border(left=thin, right=thin, top=thin, bottom=thin)

            row[4].alignment = Alignment(horizontal="center", vertical="center")
            row[4].font = Font(name="Aptos", bold=True)
            row[4].border = Border(left=thin, right=thin, top=thin, bottom=thin)

            if row[4].value == "1":
                row[4].fill = PatternFill(start_color="32C3F6", end_color="32C3F6", fill_type="solid")
            elif row[4].value == "2":
                row[4].fill = PatternFill(start_color="20FF26", end_color="20FF26", fill_type="solid")
            elif row[4].value == "3":
                row[4].fill = PatternFill(start_color="F7FF26", end_color="F7FF26", fill_type="solid")
            elif row[4].value == "4":
                row[4].fill = PatternFill(start_color="FF734D", end_color="FF734D", fill_type="solid")
            elif row[4].value == "5":
                row[4].fill = PatternFill(start_color="F30026", end_color="F30026", fill_type="solid")
            elif row[4].value == "6":
                row[4].fill = PatternFill(start_color="CC3071", end_color="CC3071", fill_type="solid")

            foundTheEnd = False
            index = 8
            while not foundTheEnd:
                cell = row[index]
                if cell.value is None or cell.value == "":
                    foundTheEnd = True
                else:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Aptos", bold=True)
                    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

                    playerName = cell.value.split(" (")[0].strip()
                    player = players[playerName.capitalize()]
                    classColor = CLASS_LIST.get(player["class"], {}).get("color", "CCCCCC")
                    cell.fill = PatternFill(start_color=classColor, end_color=classColor, fill_type="solid")

                index += 1
                if index >= len(row):
                    foundTheEnd = True


        for column in allLootSheet.columns:
            column_letter = get_column_letter(column[0].column)
            column_size = 30
            if column_letter == "B" or column_letter == "D" or column_letter == "E":
                column_size = 4.5
            elif column_letter == "A":
                column_size = 4
            elif column_letter == "C":
                column_size = 45
            
            allLootSheet.column_dimensions[column_letter].width = column_size

    if excelType == "Class Items" or excelType == "All" :
        wowClasses.sort()
        # Create item sheets
        for wowClass in wowClasses:

            classColor = "DDDDDD"
            classBgColor = "DDDDDD"
            classBgColor = CLASS_LIST.get(wowClass, {}).get("color", "DDDDDD")
            fontColor = "000000"

            hasValues = False
            classSheet = workbook.create_sheet(title=wowClass)

            sheetPlayer = {}
            
            headers = [" ", " ", " ", " ", " "]
            # Header
            for player in players.values():
                try:
                    if player["class"] == wowClass:
                        sheetPlayer[player["name"].capitalize()] = player
                        headers.append(player["name"].capitalize())
                except KeyError:
                    print(f"Warning: Player {player} does not have a class defined. Skipping.")

            headers.append(" ")

                    # Data
            classItems = []
            for itemId, item in itemList.items():
            
                itemOffNotes = item.get("itemOffNotes", "")
                # Parse itemOffNotes to extract classes and roles
                raidItemClasses = []
                raidItemRoles = []

                # Extract classes from < >
                class_matches = re.findall(r'<([^>]+)>', itemOffNotes)
                for match in class_matches:
                    raidItemClasses.append(match.strip())

                # Extract roles from [ ]
                role_matches = re.findall(r'\[([^\]]+)\]', itemOffNotes)
                for match in role_matches:
                    raidItemRoles.append(match.strip())
                
                wow_class_info = CLASS_LIST.get(wowClass, {})

                # Simplified filtering logic
                can_use_item = False

                if raidItemClasses and not raidItemRoles:
                    # Only classes specified - filter by class
                    can_use_item = wowClass in raidItemClasses

                elif raidItemRoles and not raidItemClasses:
                    # Only roles specified - filter by role
                    can_use_item = any(role in raidItemRoles for role in wow_class_info.get("roles", []))

                elif raidItemClasses and raidItemRoles:
                    # Both classes and roles specified - filter by roles + class compatibility
                    class_match = wowClass in raidItemClasses
                    wow_class_roles = wow_class_info.get("roles", [])
                    role_match = any(role in wow_class_roles for role in raidItemRoles)

                    if class_match:
                        can_use_item = True
                    else:
                        foundRoles = wow_class_roles
                        can_use_item = True
                        for otherClass in raidItemClasses:
                            other_class_info = CLASS_LIST.get(otherClass, {})
                            other_class_roles = other_class_info.get("roles", [])
                            for role in wow_class_roles:
                                if role in other_class_roles:
                                    foundRoles.remove(role)

                            if not foundRoles:
                                can_use_item = False
                                break

                if can_use_item:
                    if hasValues == False:
                        hasValues = True
                    classItems.append(item)

            print("Class " + wowClass + " has " + str(len(classItems)) + " items.")
            
            for col_num, header in enumerate(headers, start=1):
                thin = Side(border_style="thin", color="000000")

                column_letter = get_column_letter(col_num)
                cell = classSheet.cell(row=1, column=col_num, value=header)
                if cell.value != " ":
                    cell.fill = PatternFill(start_color=classColor, end_color=classColor, fill_type="solid") 
                    cell.border = Border(thin, thin, thin, thin)
                else:
                    cell.fill = PatternFill(start_color=classBgColor, end_color=classBgColor, fill_type="solid")
                cell.font = Font(name="Aptos", bold=True, color=fontColor)
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
                column_size = 16
                if col_num == 1:
                    column_size = 4
                elif col_num == 2:
                    column_size = 4.5
                elif col_num == 3:
                    column_size = 45
                elif col_num == 4:
                    column_size = 5
                elif col_num == 5:
                    column_size = 5
                elif col_num == len(sheetPlayer) + 6:
                    column_size = 4
                classSheet.column_dimensions[column_letter].width = column_size
            
            totalRows = len(classItems) + 2
            for item in classItems:
                itemData = [""]
                itemData.append(item["itemId"])
                itemData.append(item["itemName"])
                itemData.append(item["itemNotes"])
                itemData.append(item["itemTier"])
                classSheet.append(itemData)
                #if item["itemNotes"]:
                #    itemData = ["", ""]
                #    itemData.append(item["itemNotes"])
                #    totalRows += 1
                #    classSheet.append(itemData)

            
            for row in classSheet.iter_rows(min_row=2, max_row=classSheet.max_row):
                row[0].fill = PatternFill(start_color=classBgColor, end_color=classBgColor, fill_type="solid")
                row[len(sheetPlayer)+5].fill = PatternFill(start_color=classBgColor, end_color=classBgColor, fill_type="solid")
                item_id_cell = row[1]  # Assuming itemId is in the first column
                item_id = str(item_id_cell.value)

                classSheet.row_dimensions[item_id_cell.row].height = iconHeight
                if item_id is None or item_id == "":
                    current_row = item_id_cell.row
                    start_col = 3
                    end_col = len(sheetPlayer) + 5
                    classSheet.merge_cells(start_row=current_row, start_column=start_col, end_row=current_row, end_column=end_col)
                    row[2].alignment = Alignment(horizontal="left", vertical="top")
                    row[2].font = Font(name="Aptos", bold=False)
                    row[2].fill = PatternFill(start_color="FDE9D9", end_color="FDE9D9", fill_type="solid")
                    continue
                
                if item_id not in itemsIcons.keys():
                    try:
                        media_url = f'https://eu.api.blizzard.com/data/wow/media/item/{item_id}?namespace=static-classic-eu&locale=en_GB'
                        urlHeaders = {'Authorization': f'Bearer {access_token}'}
                        media_response = requests.get(media_url, headers=urlHeaders)
                        icon_url = media_response.json()['assets'][0]['value']
                        itemsIcons[item_id] = icon_url
                    except:
                        print("Error fetching media for loot item:", loot["name"])

                icon_url = itemsIcons.get(item_id)
                if icon_url:
                # Set the cell to use the IMAGE formula
                    item_id_cell.value = f'=IMAGE("{icon_url}", 2)'

                row[2].alignment = Alignment(horizontal="left", vertical="center")
                row[2].font = Font(name="Aptos", bold=True)

                if row[3].value is not None and row[3].value != "":
                    notes = row[3].value
                    row[3].value = '=IMAGE("https://render.worldofwarcraft.com/classic-eu/icons/56/inv_misc_questionmark.jpg", 2)'
                    row[3].comment = Comment(text=notes, author="")
                else:
                    classSheet.merge_cells(start_row=item_id_cell.row, start_column=3, end_row=item_id_cell.row, end_column=4)

                row[4].alignment = Alignment(horizontal="center", vertical="center")
                row[4].font = Font(name="Aptos", bold=True)

                if row[4].value == "1":
                    row[4].fill = PatternFill(start_color="32C3F6", end_color="32C3F6", fill_type="solid")
                elif row[4].value == "2":
                    row[4].fill = PatternFill(start_color="20FF26", end_color="20FF26", fill_type="solid")
                elif row[4].value == "3":
                    row[4].fill = PatternFill(start_color="F7FF26", end_color="F7FF26", fill_type="solid")
                elif row[4].value == "4":
                    row[4].fill = PatternFill(start_color="FF734D", end_color="FF734D", fill_type="solid")
                elif row[4].value == "5":
                    row[4].fill = PatternFill(start_color="F30026", end_color="F30026", fill_type="solid")
                elif row[4].value == "6":
                    row[4].fill = PatternFill(start_color="CC3071", end_color="CC3071", fill_type="solid")

                for col_num in range(5, len(sheetPlayer) + 5):
                    currRow = row[col_num]
                    currRow.alignment = Alignment(horizontal="center", vertical="center")

                    playerName = headers[col_num]
                    playerInfo = sheetPlayer[playerName]
                    row[col_num].value = "-"

                    try:
                        playerArmory = armoryList[playerName]
                    except KeyError:
                        playerArmory = []
                        armoryList[playerName] = playerArmory
                    for armoryItem in armoryList[playerName]:
                        found = False
                        if armoryItem == row[2].value:
                            found = True
                        elif row[2].value == "Head of Nefarian":
                            if armoryItem == "Master Dragonslayer's Medallion" or armoryItem == "Master Dragonslayer's Orb" or armoryItem == "Master Dragonslayer's Ring":
                                found = True
                        if found:
                            row[col_num].value = "Equipped"
                            row[col_num].fill = PatternFill(start_color="A1FB8E", end_color="A1FB8E", fill_type="solid")
                            break

                    for loot in playerInfo["loot"].values():
                        if loot["name"] == row[2].value:
                            row[col_num].value = "LC " + loot["receivedDate"]
                            row[col_num].fill = PatternFill(start_color="75F94D", end_color="75F94D", fill_type="solid")
                            break
                    
                    # Find the current item's details
                    current_item = next((item for item in classItems if item["itemId"] == item_id), None)
                    if current_item:
                        playerRole = playerInfo.get("role", "")
                        itemOffNotes = current_item.get("itemOffNotes", "")
                        
                        # Parse itemOffNotes to extract classes and roles
                        raidItemClasses = []
                        raidItemRoles = []

                        # Extract classes from < >
                        class_matches = re.findall(r'<([^>]+)>', itemOffNotes)
                        for match in class_matches:
                            raidItemClasses.append(match.strip())

                        # Extract roles from [ ]
                        role_matches = re.findall(r'\[([^\]]+)\]', itemOffNotes)
                        for match in role_matches:
                            raidItemRoles.append(match.strip())
                        
                        wow_class_info = CLASS_LIST.get(wowClass, {})

                        # Simplified filtering logic
                        can_use_item = False

                        if raidItemClasses and not raidItemRoles:
                            # Only classes specified - filter by class
                            can_use_item = wowClass in raidItemClasses

                        elif raidItemRoles and not raidItemClasses:
                            # Only roles specified - filter by role
                            can_use_item = playerRole in raidItemRoles

                        elif raidItemClasses and raidItemRoles:
                            # Both classes and roles specified - filter by roles + class compatibility
                            class_match = wowClass in raidItemClasses
                            wow_class_roles = wow_class_info.get("roles", [])
                            role_match = playerRole in raidItemRoles

                            if class_match and role_match:
                                can_use_item = True
                            elif class_match:
                                # Only class match - check if no raid item roles are shared with current class
                                wow_class_roles = wow_class_info.get("roles", [])
                                shared_roles = any(role in wow_class_roles for role in raidItemRoles)
                                can_use_item = not shared_roles
                            elif role_match:
                                # Only role match - check if no raid item classes share the role with player
                                shared_class_role = False
                                for raid_class in raidItemClasses:
                                    raid_class_info = CLASS_LIST.get(raid_class, {})
                                    raid_class_roles = raid_class_info.get("roles", [])
                                    if player_role in raid_class_roles:
                                        shared_class_role = True
                                    break
                                # Check if this player's role makes the item OS
                                can_use_item = not shared_class_role                        
                            
                        if not can_use_item:
                            row[col_num].value = "OS"
                            row[col_num].fill = PatternFill(start_color="9DC0FA", end_color="9DC0FA", fill_type="solid")

            # Remove rows where all players only have "OS" values
            rows_to_remove = []
            
            for row_idx in range(2, classSheet.max_row + 1):  # Start from row 2 (after header)
                # Check if this row has an item (has an item ID)
                item_id_cell = classSheet.cell(row=row_idx, column=2)
                if item_id_cell.value is None or item_id_cell.value == "":
                    continue  # Skip non-item rows
                
                # Check all player columns for this item
                all_os_or_empty = True
                has_any_value = False
                
                for col_num in range(6, len(sheetPlayer) + 6):  # Player columns start at column 6
                    cell_value = classSheet.cell(row=row_idx, column=col_num).value
                    
                    if cell_value and cell_value != "-":
                        has_any_value = True
                        if cell_value != "OS":
                            all_os_or_empty = False
                            break
                
                # If the row has values and they're all "OS", mark it for removal
                if has_any_value and all_os_or_empty:
                    rows_to_remove.append(row_idx)
                    item_name = classSheet.cell(row=row_idx, column=3).value
                    print(f"Removing item '{item_name}' from {wowClass} - all players have OS only")

            # Remove rows in reverse order to avoid index shifting issues
            for row_idx in reversed(rows_to_remove):
                classSheet.delete_rows(row_idx, 1)

            if rows_to_remove:
                print(f"Removed {len(rows_to_remove)} OS-only items from {wowClass} sheet")
                # Update totalRows after removing rows
                totalRows = totalRows - len(rows_to_remove)

            for col_num in range(1, len(sheetPlayer) + 7):
                finalCell = classSheet.cell(row=totalRows, column=col_num)
                finalCell.fill = PatternFill(start_color=classBgColor, end_color=classBgColor, fill_type="solid")


            # Define your data area
            min_row = 1
            max_row = classSheet.max_row
            min_col = 1
            max_col = classSheet.max_column

            thick = Side(border_style="thick", color="000000")
            thin = Side(border_style="thin", color="000000")

            for row in range(min_row, max_row + 1):
                for col in range(min_col, max_col + 1):
                    cell = classSheet.cell(row=row, column=col)

                    left_border = col == min_col + 1 and row > min_row and row < max_row
                    right_border = col == max_col - 1 and row > min_row and row < max_row
                    top_border = row == min_row + 1 and col > min_col and col < max_col
                    bottom_border = row == max_row - 1 and col > min_col and col < max_col

                    if row > min_row and row < max_row and col > min_col and col < max_col:
                        cell.border = Border(thin, thin, thin, thin)
                    
                    b = cell.border
                    border = Border(
                        left=thick if (col == min_col or left_border) else b.left,
                        right=thick if (col == max_col or right_border) else b.right,
                        top=thick if (row == min_row or top_border) else b.top,
                        bottom=thick if (row == max_row or bottom_border) else b.bottom,
                    )
                    cell.border = border

            if not hasValues:
                print(f"No items found for class {wowClass}. Skipping sheet creation.")
                workbook.remove(classSheet)
                continue
    # Item Sheets Finish

    #Save cache item icons
    with open(paths['item_icons_file'], 'w', encoding='utf-8') as f:
        json.dump(itemsIcons, f, ensure_ascii=False, indent=4)

    # Ensure we have at least one sheet in the workbook
    if len(workbook.worksheets) == 0:
        # Create a minimal info sheet if no other sheets were created
        info_sheet = workbook.create_sheet(title="Info")
        info_sheet.cell(row=1, column=1, value="No data available for the requested report type.")
        info_sheet.cell(row=2, column=1, value=f"Report type: {excelType}")
        info_sheet.cell(row=3, column=1, value="Please check your data files and try again.")

    # Return the workbook for sending to Discord
    return workbook

# Add better error handling for bot startup
if __name__ == "__main__":
    try:
        print("🤖 Starting HopiumBot...")
        logger.info("Starting HopiumBot...")
        print(f"Token present: {'Yes' if token else 'No'}")
        logger.info(f"Token present: {'Yes' if token else 'No'}")
        
        # Use INFO level for production, DEBUG for development
        log_level = logging.INFO if os.getenv('RENDER') else logging.DEBUG
        logger.info(f"Environment: {'Render (Production)' if os.getenv('RENDER') else 'Local Development'}")
        logger.info(f"Log level: {log_level}")
        
        # Start HTTP server for Render health checks (only in production)
        if os.getenv('RENDER'):
            async def health_check(request):
                return web.Response(text="HopiumBot is running!")
            
            async def start_web_server():
                app = web.Application()
                app.router.add_get('/', health_check)
                app.router.add_get('/health', health_check)
                
                port = int(os.environ.get('PORT', 8080))
                runner = web.AppRunner(app)
                await runner.setup()
                site = web.TCPSite(runner, '0.0.0.0', port)
                await site.start()
                print(f"🌐 Health check server started on port {port}")
                logger.info(f"Health check server started on port {port}")
            
            # Start web server in background
            async def main():
                await start_web_server()
                await bot.start(token)
            
            asyncio.run(main())
        else:
            # Local development - no web server needed
            logger.info("🚀 Starting bot in local development mode")
            bot.run(token)
            
    except discord.LoginFailure:
        error_msg = "Invalid bot token. Please check your DISCORD_TOKEN environment variable."
        print(f"❌ ERROR: {error_msg}")
        logger.error(error_msg)
        print("1. Go to https://discord.com/developers/applications")
        print("2. Select your application > Bot")
        print("3. Reset Token and update your environment variables")
    except discord.HTTPException as e:
        if "PHONE_REGISTRATION_ERROR" in str(e):
            error_msg = "PHONE_REGISTRATION_ERROR: This is a Discord account/token issue."
            print(f"❌ {error_msg}")
            logger.error(error_msg)
            print("Solutions:")
            print("1. Regenerate your bot token")
            print("2. Check if your Discord account needs phone verification")
            print("3. Wait 24-48 hours and try again")
        else:
            print(f"❌ HTTP Error: {e}")
            logger.error(f"HTTP Error: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        logger.error(f"Unexpected error: {e}", exc_info=True)
        # Exit gracefully in production
        import sys
        sys.exit(1)