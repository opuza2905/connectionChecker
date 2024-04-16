VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form ConnectionChecker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connection Checker - dK"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   Icon            =   "ConnectionChecker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   13815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Auto Minimized When Online"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10560
      TabIndex        =   9
      Top             =   240
      Width           =   3135
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   13815
      Begin VB.Label lblCheckLocal 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   13605
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   3
      Top             =   3360
      Width           =   13815
      Begin VB.Label lblFinalCheck 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   13605
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4605
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   12720
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   13200
      Top             =   240
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   13800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   13800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   13800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Transsion Connection Status (传音连接状态)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3120
      Width           =   4185
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Local Connection Status (本地连接状态)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Label LblCekMyIP2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   645
   End
   Begin VB.Label LblCekMyIP1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "ConnectionChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_FILENAME = &H20000

'disable close button function
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

'##################### minimize to tray
Dim nid As NOTIFYICONDATA ' trayicon variable
Sub minimize_to_tray()
Me.Hide
nid.cbSize = Len(nid)
nid.hwnd = Me.hwnd
nid.uId = vbNull
nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
nid.uCallBackMessage = WM_MOUSEMOVE
nid.hIcon = Me.Icon ' the icon will be your Form1 project icon
nid.szTip = "Connection Checker" & vbNullChar
Shell_NotifyIcon NIM_ADD, nid
End Sub
'##################### end of minimize to tray

Private Sub execCommand_2(ByVal cmd As String)
Dim result  As Long
Dim lPid    As Long
Dim lHnd    As Long
Dim lRet    As Long

cmd = "cmd /c " & cmd
result = Shell(cmd, vbHide)

lPid = result
If lPid <> 0 Then
    lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
    If lHnd <> 0 Then
        lRet = WaitForSingleObject(lHnd, INFINITE)
        CloseHandle (lHnd)
    End If
End If
Exit Sub
End Sub

''Startup Procedure
Private Sub SetStartup()
    Dim WshShell As Object
    Dim ShortcutPath As String
    Dim StartupFolder As String
    Dim AppName As String

    ' Buat objek Shell
    Set WshShell = CreateObject("WScript.Shell")

    ' Tentukan nama aplikasi Anda
    AppName = "ConnectionCheck.exe"

    ' Dapatkan path ke folder Startup
    StartupFolder = WshShell.SpecialFolders("Startup")

    ' Tentukan path lengkap untuk pintasan
    ShortcutPath = StartupFolder & "\" & AppName & ".lnk"

    ' Cek apakah pintasan sudah ada
    If Not FileExists(ShortcutPath) Then
        ' Jika pintasan belum ada, buat pintasan baru
        CreateShortcut WshShell, ShortcutPath, App.Path & "\" & AppName
    End If
End Sub

Private Function FileExists(ByVal FilePath As String) As Boolean
    FileExists = (Dir(FilePath) <> "")
End Function

Private Sub CreateShortcut(ByVal WshShell As Object, ByVal ShortcutPath As String, ByVal TargetPath As String)
    Dim Shortcut As Object

    ' Buat objek Shortcut
    Set Shortcut = WshShell.CreateShortcut(ShortcutPath)

    ' Atur properti shortcut
    Shortcut.TargetPath = TargetPath
    Shortcut.WorkingDirectory = App.Path
    Shortcut.WindowStyle = 1 ' Normal window
    Shortcut.IconLocation = TargetPath & ",0"

    ' Simpan pintasan
    Shortcut.Save
End Sub
Private Sub Form_Load()
Call SetStartup

AlwaysOnTop Me.hwnd, True
Dim hMenu As Long
hMenu = GetSystemMenu(Me.hwnd, 0)
Call RemoveMenu(hMenu, 6, MF_BYPOSITION)

Check1.Value = False
Dim serverLocal As String
Dim serverPusat As String
serverLocal = "172.17.46.1"
serverPusat = "tool-auth.transsion-os.com"

LblCekMyIP1.FontSize = 9
LblCekMyIP2.FontSize = 9
LblCekMyIP2.Caption = "Your IP : " & Winsock1.LocalIP & ""
LblCekMyIP1.Caption = "" & Winsock1.LocalHostName & ""
SB.Panels.Item(1).Text = "Local Target : " & serverLocal & ""
SB.Panels.Item(2).Text = "Transsion Address : " & serverPusat & ""
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim msg As Long
Dim sFilter As String
msg = X / Screen.TwipsPerPixelX
Select Case msg
Case WM_LBUTTONDOWN
Me.Show ' show form
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
Case WM_LBUTTONUP
Case WM_LBUTTONDBLCLK
Case WM_RBUTTONDOWN
Case WM_RBUTTONUP
Me.Show
Shell_NotifyIcon NIM_DELETE, nid
Case WM_RBUTTONDBLCLK
End Select
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
    minimize_to_tray
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid ' del tray icon
End
End Sub
Private Sub Timer1_Timer()
Static countCheck As Integer
countCheck = countCheck + 1

Set WshShell = CreateObject("WScript.Shell")
PINGLocal = Not CBool(WshShell.run("ping -n 3 172.17.46.1", 0, True))

 If PINGLocal = True Then
    
    lblCheckLocal.Caption = "Local Connection : Normal"
    Frame2.BackColor = vbGreen
    lblCheckLocal.ForeColor = vbBlack
    
    PINGFlag = Not CBool(WshShell.run("ping -n 3 tool-auth.transsion-os.com", 0, True))
    If PINGFlag = True Then
        lblFinalCheck.Caption = "Connected to tool-auth.transsion-os.com"
        Frame1.BackColor = vbGreen
        lblFinalCheck.ForeColor = vbBlack
        If Check1.Value = 1 Then
            'Me.WindowState = vbMinimized
            minimize_to_tray
        End If
    Else
resultFalse:
        lblFinalCheck.Caption = "Having Trouble to Connect to tool-auth.transsion-os.com"
        Frame1.BackColor = vbRed
        lblFinalCheck.ForeColor = vbWhite
        Me.Show
        Me.WindowState = vbNormal
        
        'write log
'        SetAttr App.Path & "\Log", vbNormal
'        Dim I As Integer
'        Dim Reg As Object
'        Set Reg = CreateObject("WScript.Shell")
'
'        Dim s As String
'
'        s = Reg.RegRead("HKLM\SYSTEM\ControlSet001\Control\ComputerName\ComputerName\ComputerName")
'
'        I = FreeFile
'        Open App.Path & "\Log\log.k43" For Append As #I
'        Print #1, " # ===================================================="
'        Print #I, " # Disconnect Log(s)"
'        Print #I, " # " & Format(Date, "dd-MMMM-YYYY") & " - " & Time$ & ""
'        Print #1, " # ===================================================="
'        Print #I, " # Using Computer " & s & ""
'        Print #I, " # IP Address " & Winsock1.LocalIP & ""
'        Print #I, " # On Connect to tool-auth.transsion-os.com"
'        Print #1, " # ===================================================="
'        Print #1, ""
'        Print #1, ""
'        Close #I
        
        'SetAttr App.Path & "\Log", vbHidden
        
        Exit Sub
    End If
 Else
    lblCheckLocal.Caption = "Local Connection : Disconnected"
    Frame2.BackColor = vbRed
    lblCheckLocal.ForeColor = vbWhite
    LblCekMyIP2.Caption = "Your IP : " & Winsock1.LocalIP & ""
    
    'write log
'    On Error Resume Next
'    SetAttr App.Path & "\Log", vbNormal
'    Dim L As Integer
'    Dim Regx As Object
'    Set Regx = CreateObject("WScript.Shell")
'
'    Dim ss As String
'
'    ss = Regx.RegRead("HKLM\SYSTEM\ControlSet001\Control\ComputerName\ComputerName\ComputerName")
'
'    L = FreeFile
'    Open App.Path & "\Log\log.k43" For Append As #L
'    Print #L, " # ===================================================="
'    Print #L, " # Disconnect Log(s)"
'    Print #L, " # " & Format(Date, "dd-MMMM-YYYY") & " - " & Time$ & ""
'    Print #L, " # ===================================================="
'    Print #L, " # Using Computer " & ss & ""
'    Print #L, " # IP Address " & Winsock1.LocalIP & ""
'    Print #L, " # On Connect local network"
'    Print #L, " # ===================================================="
'    Print #L, ""
'    Print #L, ""
'    Close #L
    
    GoTo resultFalse
    Exit Sub
 End If
End Sub
