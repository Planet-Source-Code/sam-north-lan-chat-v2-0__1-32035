VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLanChat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LaN ChAt By PeEpS"
   ClientHeight    =   4875
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6465
   Icon            =   "frmLanChat.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin LaNChAt.ShellIcon ShellIcon 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      ToolTipText     =   "LaNcHaT"
      Icon            =   "frmLanChat.frx":0442
      SysMenu         =   0   'False
   End
   Begin VB.Timer FlashTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2070
      Top             =   2385
   End
   Begin ComctlLib.Slider sldTrans 
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Top             =   4545
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   503
      _Version        =   327682
      Min             =   50
      Max             =   255
      SelStart        =   50
      TickFrequency   =   10
      Value           =   50
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Send"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5535
      MaskColor       =   &H00C0FFFF&
      TabIndex        =   1
      ToolTipText     =   "Send"
      Top             =   4095
      Width           =   945
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   0
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      ToolTipText     =   "Message to send"
      Top             =   4095
      Width           =   5490
   End
   Begin VB.TextBox txtMsgs 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4095
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Sent messages"
      Top             =   0
      Width           =   6435
   End
   Begin MSWinsockLib.Winsock connsock 
      Left            =   4560
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock listsock 
      Left            =   2925
      Top             =   2385
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   1200
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOnTop 
         Caption         =   "&Always on top"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmLanChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'LanChat
'Nice chat peer to peer chat program

'By S North (Peeps) Feb 2002
'Qs, Comments to: billbofagends@hotmail.com


'Oi you!, yes you reading this quality code, vote for me, PLEASE!
'I need to know my code is wanted to be bothered to write future
'applications ;) Thanks!


Dim DataSent As String 'data thats going to be sent
Dim UName As String    'the user name
Dim WinUName As String 'windows user name
Dim ActiveWnd As Boolean 'set depending on the window status
Dim BanTimeOutTime As String 'Reg setting
Dim BanTimeOutDate As String 'Reg setting


'<<Send Button>>
Private Sub cmdSend_Click()
If InStr(1, txtMsg.Text, ">", vbTextCompare) > 0 Then   'Check for dodgy chars
    MsgBox "Illegal character!", vbCritical + vbOKOnly, "LaNcHaT"
    Exit Sub
End If
On Error Resume Next
FlashTimer.Enabled = False
connsock.SendData UName & ":> " & txtMsg.Text 'send msg
txtMsg.Text = ""                              'reset text box
txtMsg.SetFocus                               'return the focus
End Sub

'<<Window flash timer>>
Private Sub FlashTimer_Timer()
FlashWindow Me.hWnd, 1
End Sub


Private Sub Form_Activate()
FlashTimer.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
ShellIcon.Visible = True

MakeTopMost Me.hWnd  'Make this window on top

'Get ban (if any)
BanTimeOutTime = GetSetting("LanChat", "LocalUser", "BanTimeOutTime")
BanTimeOutDate = GetSetting("LanChat", "LocalUser", "BanTimeOutDate")

'Check for ban
If Trim(BanTimeOutTime) = "" And Trim(BanTimeOutDate) = "" Then
    BanTimeOutTime = Time
    BanTimeOutDate = Date
End If

If DateValue(BanTimeOutDate) + TimeValue(BanTimeOutTime) > Now Then
    MsgBox "You have been banned until " & " " & BanTimeOutDate & " " & BanTimeOutTime & "!", vbCritical + vbOKOnly, "LaNcHaT"
    End
End If

SaveSetting "LanChat", "LocalUser", "BanTimeOutTime", Time
SaveSetting "LanChat", "LocalUser", "BanTimeOutDate", Date

Call MakeTransparent(Me.hWnd, 200) 'set semi transparent
sldTrans.Value = 200

Call WriteReg 'Try to make it start on startup

'Check to make sure only instance of app running
If App.PrevInstance = True Then
    MsgBox "Only one instance of LaNChAT can be ran at once.", vbCritical + vbSystemModal, "LaNChAt"
    End
End If

'make this app start on start up
Call WriteReg

'set the protocol, and bind
listsock.Protocol = sckUDPProtocol
listsock.Bind 1200

'set protocol and connect to broadcast
connsock.Protocol = sckUDPProtocol
connsock.RemotePort = 1200
connsock.RemoteHost = "255.255.255.255"
'connsock.RemoteHost = "localhost" 'for testing

'prompt for user name
UName = InputBox("Please enter a user name, this will be used to identify you in LaNChAt. It Must be over 3 characters long.", "LaNChAt")

'check user name
While Trim(UName) = "" Or Len(UName) <= 3 Or InStr(1, UName, "/:listusers", vbTextCompare) > 0 Or InStr(1, UName, "/:kickuser", vbTextCompare) > 0
    UName = InputBox("Please enter a user name, this will be used to identify you in LaNChAt. It Must be over 3 characters long.", "LaNChAt")
Wend

WinUName = Environ$("USERNAME")
UName = "[" & WinUName & "] " & UName

If txtMsg.Text = "" Then
    cmdSend.Enabled = False
Else
    cmdSend.Enabled = True
End If

'notify others of join
connsock.SendData "LaNChAt UsEr JoIn::>> " & UName
txtMsg.SetFocus
End Sub


Private Sub Form_Paint()
FlashTimer.Enabled = False
End Sub

'shut down the sockets
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
FlashTimer.Enabled = False
connsock.SendData "LaNChAt UsEr LeAvE::>> " & UName 'send leave
connsock.Close                                      'close down
listsock.Close
ShellIcon.Visible = False
Unload frmAbout
End
End Sub

'accept any incoming connections
Private Sub listsock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
listsock.Accept requestID
End Sub

Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 'show about dlg
End Sub

'Exit app
Private Sub mnuExit_Click()
On Error Resume Next
Call Form_Unload(0)
End Sub

Private Sub mnuOnTop_Click()
If mnuOnTop.Checked = False Then
    mnuOnTop.Checked = True
    MakeTopMost Me.hWnd
Else
    mnuOnTop.Checked = False
    MakeNormal Me.hWnd
End If
End Sub

Private Sub shellicon_MouseDown(Button As Integer)
PopupMenu mnuFile
End Sub

Private Sub sldTrans_Change()
On Error Resume Next
Call MakeTransparent(Me.hWnd, sldTrans.Value) 'set transparent level
End Sub

'enable/disbale send button
Private Sub txtMsg_Change()
On Error Resume Next
FlashTimer.Enabled = False
If Trim(txtMsg.Text) = "" Then
    cmdSend.Enabled = False
Else
    cmdSend.Enabled = True
End If
End Sub

Private Sub listsock_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
listsock.GetData DataSent, , bytesTotal 'get the data

If Not Me.WindowState = 0 Then FlashTimer.Enabled = True

If InStr(1, DataSent, "/:listusers", vbTextCompare) > 0 Then
    connsock.SendData UName & " Machine: " & listsock.LocalHostName
    Exit Sub
End If

If InStr(1, DataSent, "/:kickuser", vbTextCompare) > 0 Then
    If InStr(1, DataSent, "/:kickuser [" & WinUName & "]", vbTextCompare) > 0 Then
        connsock.SendData "UsEr BanNeD::>>" & UName
        connsock.Close
        listsock.Close
        SaveSetting "LanChat", "LocalUser", "BanTimeOutTime", Time + TimeValue("0:10:0")
        SaveSetting "LanChat", "LocalUser", "BanTimeOutDate", Date
        MsgBox "You have been banned for 10 minutes, so go play hide and go fuck yourself!", vbCritical + vbOKOnly, "LaNcHaT"
        Call Form_Unload(0)
    End If
    Exit Sub
End If

If txtMsgs.Text <> "" Then
    txtMsgs.Text = txtMsgs.Text & vbNewLine & DataSent
End If
If txtMsgs.Text = "" Then
    txtMsgs.Text = DataSent
End If
End Sub

Private Sub connsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Err.Raise Number, Source, Description, HelpFile, HelpContext
End Sub

Private Sub listsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Err.Raise Number, Source, Description, HelpFile, HelpContext
End Sub


Private Sub txtMsg_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Shift = 1 Then
  txtMsg.PasswordChar = "*"
Else
  txtMsg.PasswordChar = ""
End If
End Sub

Private Sub txtMsgs_Change()
txtMsgs.SelStart = Len(txtMsgs.Text)
End Sub

