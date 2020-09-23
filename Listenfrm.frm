VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ListenFrm 
   Caption         =   "PiK Soft Net Tools - Listener"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   Icon            =   "ListenFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbPort 
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Text            =   "1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Listen"
      Default         =   -1  'True
      Height          =   255
      Left            =   3360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDisconnect 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton optTCP 
      Caption         =   "TCP/IP"
      Height          =   195
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.OptionButton optUDP 
      Caption         =   "UDP"
      Height          =   195
      Left            =   2280
      TabIndex        =   4
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Protocol:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Port:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
      Begin VB.TextBox txtStatus 
         Height          =   2535
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   4455
      End
   End
   Begin MSWinsockLib.Winsock ws1 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "Port:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   720
      TabIndex        =   9
      Top             =   480
      Width           =   405
   End
End
Attribute VB_Name = "ListenFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdConnect_Click()
cmdConnect.Enabled = False
cmbPort.Enabled = False
cmdDisconnect.Enabled = True
txtStatus = ""
If optTCP = True Then
    ws1.Protocol = sckTCPProtocol
End If
If optUDP = True Then
    ws1.Protocol = sckUDPProtocol
End If
On Error GoTo PortIsOpen
ws1.Close
ws1.LocalPort = cmbPort.Text
ws1.Listen
Exit Sub
PortIsOpen:
ws1.Close
If Err.Number = 10048 Then
    txtStatus = "The port " & cmbPort.Text & " is already open."
Else
    txtStatus = "Error: " & Err.Number & vbCrLf & "   " & Err.Description
End If
cmdDisconnect.Enabled = False
cmbPort.Enabled = True
cmdConnect.Enabled = True
End Sub

Private Sub cmdDisconnect_Click()
ws1.Close
cmdDisconnect.Enabled = False
cmbPort.Enabled = True
cmdConnect.Enabled = True
End Sub


Private Sub Form_Load()
optTCP = True
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4935
Me.Width = 5055
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub ws1_ConnectionRequest(ByVal requestID As Long)
 If ws1.State <> sckClosed Then ws1.Close
 ws1.Accept (requestID)
 txtStatus.Text = "Connection..."
End Sub

Private Sub ws1_DataArrival(ByVal bytesTotal As Long)
Dim strData As String
ws1.GetData strData
txtStatus.Text = txtStatus.Text & vbCrLf & " - " & strData
End Sub

Private Sub ws1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
txtStatus = "Winsock Error: " & Number & vbCrLf & "   " & descriptoin
End Sub
