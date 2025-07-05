import discord
from discord.ext import commands
import logging
from dotenv import load_dotenv
import os
import requests
import asyncio
import aiohttp

load_dotenv()
token = os.getenv('DISCORD_TOKEN')

handler = logging.FileHandler(filename='hopiumbot.log', encoding='utf-8', mode='w')
intents = discord.Intents.default()
intents.message_content = True  # Enable message content intent
intents.guilds = True  # Enable guild intents
intents.members = True  # Enable member intents

bot = commands.Bot(command_prefix='!', intents=intents)

role = "Trial"

# Store ongoing applications
active_applications = {}

async def validate_character_exists(character_name):
    """
    Check if character exists on Classic WoW Armory (for review links)
    Returns: (exists, error_message)
    """
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
    """
    Get mentions for Karumenta and Hokkies if they're in the server
    Returns: formatted mention string
    """
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
    """
    Validate character name by checking Classic WoW Armory
    Returns: (is_valid, error_message)
    """
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

# Application questions
APPLICATION_QUESTIONS = [
    "Character name:",
    "Class/Spec:",
    "What country are you from and how old are you?",
    "Please tell us a bit about yourself, who are you outside of the game?",
    "Explain your WoW experience. Include logs of past relevant characters (Classic/SoM//SoD/Retail).",
    "We require a few things from every raider in the guild. To have above average performance for your class and atleast 80% raid attendance. Can you fulfill these requirements?",
    "Why did you choose to apply to <Hopium>?",
    "Can someone in <Hopium> vouch for you?",
    "Surprise us! What's something you'd like to tell us, it can be absolutely anything!"
]

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
                'guild_id': interaction.guild.id
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
                    value="Please create a character on https://thatsmybis.com/ and add him to the guild from the home page. Once you do it notify us and feel free to create a wishlist for the current phase ( No Ony/ZG ). Thanks!",
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
    
    # Special validation for character name (first question)
    if app_data['question_index'] == 0:
        character_name = message.content.strip()
        
        # Validate character name
        guild = bot.get_guild(app_data['guild_id'])
        is_valid, error_message = await validate_character_name(character_name, guild)
        
        if not is_valid:
            # Send error message and ask for character name again
            error_embed = discord.Embed(
                title="❌ Character Validation Failed",
                description=error_message,
                color=0xff0000
            )
            error_embed.add_field(
                name=f"Question 1/{len(APPLICATION_QUESTIONS)}",
                value=APPLICATION_QUESTIONS[0],
                inline=False
            )
            error_embed.set_footer(text="Please provide the correct character name or type 'cancel' to cancel the application.")
            await message.channel.send(embed=error_embed)
            return  # Don't advance to next question, ask again
    
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
async def on_ready():
    print(f'Logged in as {bot.user.name} - {bot.user.id}')
    print(f'Bot is in {len(bot.guilds)} guilds')
    print('------')

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

# Validate token before starting
if not token:
    print("❌ ERROR: No Discord token found in .env file!")
    print("Please create a .env file with: DISCORD_TOKEN=your_bot_token_here")
    exit(1)

@bot.command()
async def setupHopium(ctx):
    guild = ctx.guild
    
    # Check if "Recruitment" category exists
    recruitment_category = discord.utils.get(guild.categories, name="Recruitment")
    
    if not recruitment_category:
        # Create the "Recruitment" category with restricted permissions
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        recruitment_category = await guild.create_category("Recruitment", overwrites=overwrites)
        await ctx.send("Created 'Recruitment' category!")
    
    # Check if "apply-here" channel exists in the category
    apply_channel = discord.utils.get(recruitment_category.channels, name="✍apply-here")
    
    if not apply_channel:
        # Create the "apply-here" channel in the Recruitment category with restricted permissions
        overwrites = {
            guild.default_role: discord.PermissionOverwrite(read_messages=False, send_messages=False, view_channel=False)
        }
        apply_channel = await guild.create_text_channel("✍apply-here", category=recruitment_category, overwrites=overwrites)
        await ctx.send("Created 'apply-here' channel in the Recruitment category!")
    
    # Clear the apply-here channel to ensure clean state
    await apply_channel.purge()
    await ctx.send("Cleared 'apply-here' channel!")
    
    # Create the application message with button
    embed = discord.Embed(
        title="📋 Application for Hopium Guild",
        description=f"Click the button below to start your application process!\nIf anything goes wrong, please contact {await get_staff_mentions(guild)}.",
        color=0x00ff00
    )
    
    view = ApplicationView()
    await apply_channel.send(embed=embed, view=view)
    await ctx.send("Application message sent with Apply button!")
    
    await ctx.send('Setup completed!')

# Add better error handling for bot startup
if __name__ == "__main__":
    try:
        print("🤖 Starting HopiumBot...")
        print(f"Token present: {'Yes' if token else 'No'}")
        
        # Use INFO level for production, DEBUG for development
        log_level = logging.INFO if os.getenv('RENDER') else logging.DEBUG
        
        bot.run(token, log_handler=handler, log_level=log_level)
    except discord.LoginFailure:
        print("❌ ERROR: Invalid bot token. Please check your DISCORD_TOKEN environment variable.")
        print("1. Go to https://discord.com/developers/applications")
        print("2. Select your application > Bot")
        print("3. Reset Token and update your environment variables")
    except discord.HTTPException as e:
        if "PHONE_REGISTRATION_ERROR" in str(e):
            print("❌ PHONE_REGISTRATION_ERROR: This is a Discord account/token issue.")
            print("Solutions:")
            print("1. Regenerate your bot token")
            print("2. Check if your Discord account needs phone verification")
            print("3. Wait 24-48 hours and try again")
        else:
            print(f"❌ HTTP Error: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        # Exit gracefully in production
        import sys
        sys.exit(1)