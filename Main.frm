VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Main 
   BorderStyle     =   0  'None
   Caption         =   "PiK Soft Net Tools"
   ClientHeight    =   4500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6750
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":0442
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5040
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton RegIt 
      Caption         =   "Register"
      Height          =   255
      Left            =   480
      TabIndex        =   13
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Speed 
      Caption         =   "Speed Check"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Winsck 
      Caption         =   "About"
      Height          =   255
      Left            =   5040
      TabIndex        =   11
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton FingerBtn 
      Caption         =   "Finger"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton ListenBtn 
      Caption         =   "Listener"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton WhoisBtn 
      Caption         =   "Whois"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Source 
      Caption         =   "Get HTML"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Mail 
      Caption         =   "Email"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Trace 
      Caption         =   "TraceRoute"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Lookup 
      Caption         =   "Host Lookup"
      Enabled         =   0   'False
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Scanner 
      Caption         =   "Port Scan"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Ping 
      Caption         =   "Ping"
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton ConnectBtn 
      Caption         =   "Raw Connect"
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox HostTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox IPTxt 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   960
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Label Registered 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UNREGISTERED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   18
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   8
      X2              =   440
      Y1              =   224
      Y2              =   224
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   224
      X2              =   224
      Y1              =   96
      Y2              =   224
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8
      X2              =   440
      Y1              =   96
      Y2              =   96
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Host:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label IPLbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Local IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ConnectBtn_Click()
Connect.Show
End Sub

Private Sub Exit_Click()
Unload Me
End
End Sub

Private Sub FingerBtn_Click()
FingerFrm.Show
End Sub

Private Sub Form_DblClick()
Me.WindowState = vbMinimized
End Sub

Private Sub Form_Load()
Dim NameStr As String, SerStr As String
IPTxt.Text = Winsock1.LocalIP
Hosttxt.Text = Winsock1.LocalHostName
NameStr = GetSetting("PiK Soft Net Tools", "Info", "Name", , HKEY_CURRENT_USER)

Register.NameTxt.Text = GetSetting("PiK Soft Net Tools", "Info", "Name", , HKEY_CURRENT_USER)
Register.Email.Text = GetSetting("PiK Soft Net Tools", "Info", "Email", , HKEY_CURRENT_USER)
Register.Serial.Text = GetSetting("PiK Soft Net Tools", "Info", "Regcode", , HKEY_CURRENT_USER)
SerStr = GetSetting("PiK Soft Net Tools", "Info", "Regcode", , HKEY_CURRENT_USER)

If Len(SerStr) > 0 Then
If Serial_Check = SerStr Then
Lookup.Enabled = True
ListenBtn.Enabled = True
Speed.Enabled = True
Scanner.Enabled = True
Trace.Enabled = True
FingerBtn.Enabled = True
RegIt.Visible = False
Registered.Left = 32
Registered.Alignment = vbLeftJustify
Registered.Caption = "Welcome: " & NameStr
Aboutfrm.Reglbl.Caption = "This Product is Registered to:" & vbCrLf & NameStr
End If
End If

Serial_Check
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub ListenBtn_Click()
ListenFrm.Show
End Sub

Private Sub Lookup_Click()
LookupFrm.Show
End Sub

Private Sub Mail_Click()
EmailFrm.Show
End Sub

Private Sub Ping_Click()
Pingfrm.Show
End Sub

Private Sub RegIt_Click()
Register.Show
End Sub

Private Sub Scanner_Click()
ScannerFrm.lblCurrent = 0
ScannerFrm.Show
End Sub

Private Sub Source_Click()
GetHTML.Show
End Sub

Private Sub Speed_Click()
SpeedChk.Show
End Sub

Private Sub Trace_Click()
Tracefrm.Show
End Sub

Private Sub WhoisBtn_Click()
WhoisFrm.Show
End Sub

Private Sub Winsck_Click()
Aboutfrm.Show 1
End Sub
