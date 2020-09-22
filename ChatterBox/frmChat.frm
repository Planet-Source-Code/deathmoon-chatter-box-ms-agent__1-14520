VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   Caption         =   "Chatter Box"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6450
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVolume 
      Caption         =   "&Volume"
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Disconnect"
      Height          =   375
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   4920
      TabIndex        =   21
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Listen"
      Height          =   375
      Left            =   4920
      TabIndex        =   20
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Frame fraGeneralInfo 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   4680
      TabIndex        =   11
      Top             =   120
      Width           =   1695
      Begin VB.TextBox txtPort 
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Text            =   "2701"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtLocal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "127.0.0.1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtRemote 
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Text            =   "0.0.0.0"
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtNick 
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblPort 
         Caption         =   "Port"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "If you are the host, this is the port others will connect to you on."
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblLocalHost 
         Caption         =   "Local Host"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         ToolTipText     =   "This is your IP Address."
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblRemoteHost 
         Caption         =   "Remote Host"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "This is the IP Address / Computer Name you want to connect to."
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblNickName 
         Caption         =   "Nick Name"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin MSWinsockLib.Winsock wData 
      Left            =   6480
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cboAgents 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdColor 
      Caption         =   "&Color"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6480
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdUnderline 
      Caption         =   "Underline"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdItalic 
      Caption         =   "&Italic"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdBold 
      Caption         =   "&Bold"
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin RichTextLib.RichTextBox rtbMsgToSend 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":08CA
   End
   Begin RichTextLib.RichTextBox rtbDialog 
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4048
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":0984
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   4455
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   6480
      Top             =   120
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnuHelpSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sAgentFiles As String
Public sSelectedAgent As String
Public sLoadedAgent As String
Public sSpeak As String

Public Sub LoadAgentCharacter(sAgent As String)

    sAgentFiles = sAgentPath & "\" & sAgent & ".acs"
On Error GoTo Trap:
    Agent1.Characters.Load sAgent, sAgentFiles
On Error GoTo LocalTrap:

    If sLoadedAgent = "" Then
        sLoadedAgent = sAgent
        Set cAgent = Agent1.Characters(sAgent)
        cAgent.LanguageID = &H409  ' ENGLISH
        cAgent.Show
        cAgent.Speak "Welcome..."
        basMain.MovementsForCharacters sAgent, cAgent
        Exit Sub
    ElseIf sLoadedAgent <> "" Then
        'This is if a charater has already been loaded
        'it will hide him and open the new one.
        cAgent.Speak "So that is how you are..."
        cAgent.Hide
        Wait (5)
        '--Unload the previous character
        Agent1.Characters.Unload sLoadedAgent
        Wait (2)
        Set cAgent = Nothing
        Set cAgent = Agent1.Characters(sAgent)
        cAgent.LanguageID = &H409  ' ENGLISH
        cAgent.Show
        cAgent.Speak "I'm glad you got rid of that other guy"
        sLoadedAgent = sAgent
        basMain.MovementsForCharacters sAgent, cAgent
        Exit Sub
    End If
    '
Trap:
    Dim obj As New clsWinDir
    Dim sTmp As String
    sTmp = obj.Get_Windows_Directory
    sAgentFiles = sTmp & "\" & sAgent & ".acs"
    Agent1.Characters.Load sAgent, sAgentFiles
    Resume Next
LocalTrap:
    ' Try looking for the file in one more "standard" location before giving error
    MsgBox "You do not have this Agent!  Please download!", vbCritical + vbOKOnly, "Error"
    Exit Sub
End Sub

Public Sub UnloadAgentCharacter()
    Me.WindowState = 1
    Me.Hide
    If sSelectedAgent <> "" Then
        cAgent.Speak "Bye Bye..."
        Wait (1)
        cAgent.Hide
        basMain.Wait (1)
        Agent1.Characters.Unload sLoadedAgent
        Set cAgent = Nothing
    End If
    Unload Me
    End
End Sub

Private Sub cboAgents_Click()
    If sLoadedAgent = Me.cboAgents.Text Then
        Exit Sub
    End If
    sSelectedAgent = Me.cboAgents.Text
    Me.LoadAgentCharacter sSelectedAgent
End Sub

Private Sub cmdColor_Click()
    Me.CommonDialog1.ShowColor
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdListen_Click()
    If Me.txtPort.Text = "" Then
        MsgBox "You must have a port listed before listening...", vbOKOnly + vbInformation, "Error"
        Exit Sub
    End If
    On Error Resume Next
    'Set the port to listen to:
    wData.LocalPort = txtPort.Text
    'Start listening for a connection:
    wData.Listen
    'Inform the user we are listening for a connection:
    lblStatus.Caption = "Listening..."
    'If there was an error, inform the user:
    If Err Then lblStatus.Caption = Err.Description
End Sub

Private Sub cmdDisconnect_Click()
    'Disconnect the current connection:
    wData.Close
    'Hide the disconnect button since we are done disconnecting:
    cmdDisconnect.Visible = False
End Sub

Private Sub cmdConnect_Click()
    On Error Resume Next
    'Close the current connection for new connection:
    wData.Close
    'Connect to the remote computer
    wData.Connect txtRemote.Text, txtPort.Text
    'Tell the user we are connecting:
    lblStatus.Caption = "Connecting..."
    'Show the dsconnect button since we are connecting:
    cmdDisconnect.Visible = True
    'If there was an error, inform the user:
    If Err Then lblStatus.Caption = Err.Description
End Sub

Private Sub cmdSend_Click()
    Dim SendStr As String
    On Error Resume Next
    'Put what we are about so send into a variable:
    SendStr = txtNick & ":" & vbTab & rtbMsgToSend.Text
    'Send the message:
    wData.SendData SendStr
    'Put the current selection to the end:
    rtbDialog.SelStart = Len(rtbDialog.Text)
    'Set the color to blue:
    rtbDialog.SelColor = vbBlue
    'Set the text to our nickname:
    rtbDialog.SelText = txtNick & ":" & vbTab
    'Set selection to the end:
    rtbDialog.SelStart = Len(rtbDialog.Text)
    'Change color back to black:
    rtbDialog.SelColor = vbBlack
    'Set the text to the message we sent:
    rtbDialog.SelText = rtbMsgToSend.Text & vbCrLf
    'Clear Message To Send Text Box
    Me.rtbMsgToSend.Text = ""
    'If there was an error, inform the user:
    If Err Then lblStatus.Caption = Err.Description
End Sub

Private Sub cmdVolume_Click()
On Error Resume Next
    Shell ("sndvol32.exe"), vbNormalFocus
End Sub

Private Sub Form_Load()
    Me.Icon = LoadResPicture(101, vbResIcon)
    
    basMain.GetUsrSettings
        Me.txtNick.Text = basMain.sNickName
        Me.txtRemote.Text = basMain.sRemoteHost
        Me.txtPort.Text = basMain.sPort
    FillCombo cboAgents
    '// gets ip address from winsock
    txtLocal.Text = wData.LocalIP
    Me.cboAgents.Text = "Merlin"
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    '
    Me.Width = 6600
    Me.Height = 4980
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iResponse As Integer
    iResponse = MsgBox("Are you sure you want to exit?", vbQuestion + vbYesNo, "Exit Agent Chat")
    If iResponse = vbYes Then
        cmdDisconnect_Click
        Me.UnloadAgentCharacter
        basMain.SaveUsrSettings Me.txtNick.Text, _
            Me.txtRemote.Text, Me.txtPort.Text
        End
    Else
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    frmSplash.Show
End Sub

Private Sub mnuHelpHelp_Click()
    MsgBox "Nothing yet...email me if you have problems! (deathmoon91@yahoo.com)", vbOKOnly
End Sub

Private Sub rtbMsgToSend_GotFocus()
    Me.cmdSend.Default = True
End Sub

Private Sub rtbMsgToSend_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         cmdSend_Click
    End If
End Sub

Private Sub rtbMsgToSend_LostFocus()
    Me.cmdSend.Default = False
End Sub

Private Sub wData_Close()
    'Inform the user that we have closed the connection:
    lblStatus.Caption = "Connection Closed"
End Sub

Private Sub wData_Connect()
    'Inform the user we have connected:
    lblStatus.Caption = "Connected!"
End Sub

Private Sub wData_ConnectionRequest(ByVal requestID As Long)
    'Close the current connection:
    wData.Close
    'Accept the connection request:
    wData.Accept requestID
    'Inform the user we have accepted a connection:
    lblStatus.Caption = "Connection Accepted!"
End Sub

Private Sub wData_DataArrival(ByVal bytesTotal As Long)
    Dim nData As String
    On Error Resume Next
    'Get the data being sent to us:
    wData.GetData nData
    'Set the selection to the end:
    rtbDialog.SelStart = Len(rtbDialog.Text)
    'Set color to red:
    rtbDialog.SelColor = vbRed
    'Put the nickname of the other person:
    rtbDialog.SelText = Left(nData, InStr(1, nData, ":"))
    
    'GET THE TEXT THE OTHER PERSON SENT
    Dim x As Integer
    Dim l As Integer
    Dim y As Integer
    l = Len(nData)
    x = InStr(1, nData, ":")    ' This gets the # of spaces from their username to the :
    y = l - x
    sSpeak = Right(nData, y)
    
    'Set the selection to the end:
    rtbDialog.SelStart = Len(rtbDialog.Text)
    'Change the color back to black:
    rtbDialog.SelColor = vbBlack
    'Set the text to the message we received:
    rtbDialog.SelText = Mid(nData, InStr(1, nData, ":") + 1) & vbCrLf
    cAgent.Speak sSpeak
    'If there was an error, inform the user:
    If Err Then lblStatus.Caption = Err.Description
End Sub

Private Sub wData_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Inform the user of an error:
    lblStatus.Caption = Description
End Sub
