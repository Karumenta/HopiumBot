# HopiumBot - Discord Guild Application Bot

A Discord bot for managing guild applications with automated character validation, role management, and application review system.

## Features

- **Application System**: Interactive DM-based application process
- **Character Validation**: Validates characters using Classic WoW Armory
- **Role Management**: Automatically manages Trial, Raider, Officer, and Guild Leader roles
- **Dynamic Staff Mentions**: Automatically mentions available staff members
- **Channel Management**: Creates and manages application/review channels with proper permissions
- **Screenshot Integration**: Optional character screenshots via API services

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Set up environment variables:
   - `DISCORD_TOKEN`: Your Discord bot token
   - `HCTI_API_USER_ID`: (Optional) Screenshot API user ID
   - `HCTI_API_KEY`: (Optional) Screenshot API key

3. Run the bot:
   ```bash
   python main.py
   ```

4. Set up the application system:
   ```
   !setupHopium
   ```

## Deployment

This bot is configured for deployment on platforms like Render, Railway, or Fly.io.

## Commands

- `!setupHopium`: Initialize the application system and create necessary channels

## Environment Variables Required

- `DISCORD_TOKEN`: Your Discord bot token (required)
- `HCTI_API_USER_ID`: Screenshot service user ID (optional)
- `HCTI_API_KEY`: Screenshot service API key (optional)
