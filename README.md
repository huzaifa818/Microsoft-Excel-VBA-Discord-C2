# Microsoft-Excel VBA Discord C2 (Command & Control)

A proof-of-concept Command & Control (C2) framework built entirely in Microsoft Excel VBA. This script turns an Excel workbook into a remote administration tool that polls a Discord channel for commands, executes them on a Windows machine (e.g., a VirtualBox VM), and returns the output to Discord.

**⚠️ DISCLAIMER:** This project is for **educational purposes only**. It demonstrates API integration, WMI queries, and shell execution in VBA. Do not use this on machines you do not own.

## 🚀 Features
- **Asynchronous Polling:** Checks Discord for new messages every 5 seconds without freezing Excel.
- **Command Execution:** Run CMD commands, launch executables.
- **File Transfer:** Download files from Discord to the VM (`!upload`), and upload files from the VM to Discord (`!download`).
- **System Recon:** Dump OS, CPU, RAM, BIOS, and Network info (`!sysinfo`).
- **Geolocation:** Fetch public IP location data (`!location`).
- **Screen Capture:** Take screenshots and upload them directly to Discord.
- **Utility:** Read clipboard, change directory, list files, lock workstation.

## 🛠️ Prerequisites
1. Microsoft Excel (Desktop version).
2. A Discord Server and a Discord Bot Token.
3. **Microsoft Scripting Runtime** reference enabled in VBA.

## 📦 Installation Guide

### 1. Discord Bot Setup
1. Go to the [Discord Developer Portal](https://discord.com/developers/applications).
2. Create an App, go to **Bot**, and copy the Token.
3. Enable **Message Content Intent** under Privileged Gateway Intents.
4. Invite the bot to your server with permissions: `Send Messages`, `Read Message History`, `Attach Files`, `Add Reactions`.
5. Enable Developer Mode in Discord settings. Right-click your text channel -> **Copy Channel ID**. Right-click your name -> **Copy User ID**.

### 2. Excel Setup
1. Open Excel, save a new file as **Excel Macro-Enabled Workbook (*.xlsm)**.
2. Press `Alt + F11` to open the VBA Editor.
3. Go to `Tools > References` and check **Microsoft Scripting Runtime**.
4. Download `JsonConverter.bas` from the VBA-JSON repo and import it via `File > Import File`.

### 3. Inserting the Code
Create the following files in the VBA Editor and paste the code provided below.

- **`ThisWorkbook`** (Excel Object)
- **`DiscordBot`** (Class Module - *Ensure `Instancing` is set to `2 - PublicNotCreatable`*)
- **`DiscordApi`** (Standard Module)
- **`BotCommands`** (Standard Module)

### 4. Configuration
Update the `DiscordApi` module with your specific IDs and Token:
```vba
Public Const DISCORD_CHANNEL_ID   As String = "YOUR_CHANNEL_ID_HERE"
Public Const DISCORD_ALLOWED_USER_ID As String = "YOUR_USER_ID_HERE"

Public Function DiscordToken() As String
    DiscordToken = "YOUR_BOT_TOKEN_HERE"
End Function
