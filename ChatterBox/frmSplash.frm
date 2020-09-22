VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   7185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   3150
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7140
      Begin VB.Timer Timer1 
         Interval        =   3500
         Left            =   4560
         Top             =   1920
      End
      Begin VB.Label lblVersion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   9
         Top             =   675
         Width           =   1935
      End
      Begin VB.Label lblUrl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "www.planet-source-code.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4320
         TabIndex        =   8
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   1485
         Left            =   240
         Picture         =   "frmSplash.frx":0000
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblAppName 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   270
         TabIndex        =   7
         Top             =   675
         Width           =   1665
      End
      Begin VB.Label lblFAX 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   4
         Left            =   4320
         TabIndex        =   6
         Top             =   1200
         Width           =   2625
      End
      Begin VB.Label lblEMail 
         BackColor       =   &H00C0C0C0&
         Caption         =   "email: deathmoon91@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   3
         Left            =   4320
         MousePointer    =   4  'Icon
         TabIndex        =   5
         Top             =   1440
         Width           =   2850
      End
      Begin VB.Label lblCityS 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   1
         Left            =   4320
         TabIndex        =   4
         Top             =   930
         Width           =   2625
      End
      Begin VB.Label lblPO 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   0
         Left            =   4320
         TabIndex        =   3
         Top             =   705
         Width           =   2625
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Chatter Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   270
         TabIndex        =   2
         Tag             =   "CompanyProduct"
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label lblWarning 
         BackColor       =   &H00C0C0C0&
         Caption         =   $"frmSplash.frx":24BE
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         TabIndex        =   1
         Tag             =   "Warning"
         Top             =   2400
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Image1.Picture = LoadResPicture(101, vbResBitmap)
    Me.lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    frmChat.Show
End Sub

Private Sub fraMainFrame_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub Label2_Click()
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()
    Unload Me
End Sub

Private Sub lblProductName_Click(Index As Integer)
    Unload Me
End Sub

Private Sub lblEMail_Click(Index As Integer)
    Shell ("explorer mailto:deathmoon91@yahoo.com"), vbNormalNoFocus
End Sub

Private Sub lblUrl_Click()
    Shell ("explorer http://www.planet-source-code.com"), vbNormalNoFocus
End Sub

Private Sub lblVersion_Click()
    Unload Me
End Sub

Private Sub lblWarning_Click()
    Unload Me
End Sub

Private Sub picLogo_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
