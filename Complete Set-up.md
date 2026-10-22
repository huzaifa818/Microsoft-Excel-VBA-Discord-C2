
# 🚀 The Complete Step-by-Step Guide

## Phase 1: Discord Setup
Before touching Excel, we need to create the bot and get the required IDs.
Check this repo for setup discord bot https://github.com/huzaifa818/Discord-c2 
or go to following steps:
1. **Create the App:** Go to the [Discord Developer Portal](https://discord.com/developers/applications) and click **New Application**. Name it and accept the terms.
2. **Get the Token:** Click **Bot** on the left menu. Click **Reset Token**, copy it, and save it somewhere safe. *(Never share this token!)*
3. **Enable Intents:** Scroll down on the Bot page to **Privileged Gateway Intents**. Turn **ON** the `MESSAGE CONTENT INTENT`. Click **Save Changes**.
4. **Invite the Bot:** Go to **OAuth2 > URL Generator**. Check the `bot` scope. Under Bot Permissions, check `View Channels`, `Send Messages`, `Read Message History`, `Attach Files`, and `Add Reactions`. Copy the Generated URL, paste it into your browser, and invite the bot to your server.
5. **Get IDs:** Open Discord, go to **User Settings > Advanced**, and turn **Developer Mode ON**. 
   * Right-click your server's text channel and click **Copy Channel ID**.
   * Right-click your own profile picture and click **Copy User ID**.

## Phase 2: Excel Preparation
1. Download this file `DiscordBot.xlsm`
2. Open Microsoft Excel (Desktop version). 
3. **Open VBA:** Press `Alt + F11` to open the VBA Editor. (make sure you enable developer option and macros)
4. **Add Reference:** In the VBA Editor, click `Tools > References`. Scroll down, check the box for **Microsoft Scripting Runtime**, and click OK.

<img width="956" height="502" alt="Screenshot 2026-10-23 013759" src="https://github.com/user-attachments/assets/65065113-120e-4134-b3b6-c9c97805469b" />

## Phase 3: Configuration & Testing
1. Go to the **`DiscordApi`** module. Replace the placeholders at the top of the file:
   * `"YOUR_CHANNEL_ID_HERE"` -> Your copied Channel ID.
   * `"YOUR_USER_ID_HERE"` -> Your copied User ID.
   * `"YOUR_BOT_TOKEN_HERE"` -> Your Bot Token.
     <img width="952" height="491" alt="Screenshot 2026-10-23 014740" src="https://github.com/user-attachments/assets/e3c6cf39-dfdb-447a-9e06-854c36091fab" />

2. Click **Debug > Compile VBAProject** to ensure there are no typos. If it succeeds, close the VBA editor.
3. Save the Excel file.
4. **Close Excel completely and reopen it.**
5. If a yellow bar appears, click **Enable Content**.
6. Press `Alt + F11`, open the **Immediate Window** (`Ctrl + G`). You should see `Bot listening. Baseline ID: ...`
7. Go to Discord, type **`!ping`**.

***

## 💻 Command List
- `!help` - Show all commands
- `!ping` - Check bot latency
- `!sysinfo` - Full VM system information
- `!location` - Public IP geolocation
- `!ip` - Local IP configuration
- `!uptime` - Boot time and uptime
- `!ps` - Running processes
- `!pwd` / `!cd` / `!ls` / `!cat` - File system navigation
- `!cmd <command>` - Run a shell command
- `!run <path>` - Execute a program on the VM
- `!upload [path]` - Upload a file from Discord to VM (drag & drop file into Discord)
- `!download <path>` - Send a file from VM to Discord
- `!screenshot` - Capture screen and upload
- `!clip` - Read VM clipboard
- `!msgbox <text>` - Popup a box on the VM
- `!volume <0-100>` - Adjust volume
- `!wallpaper <url>` - Change desktop background
- `!lock` - Lock the workstation
```
```
