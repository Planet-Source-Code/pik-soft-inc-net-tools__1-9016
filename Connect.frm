VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Connect 
   Caption         =   "PiK Soft Net Tools - Connect"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5280
   Icon            =   "Connect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   5055
      Begin VB.TextBox CodeWin 
         Height          =   1215
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label4 
         Caption         =   "Data Recieved:"
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
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   4080
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton Connect 
         Caption         =   "Connect"
         Height          =   255
         Left            =   4080
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Send 
         Caption         =   "Send"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   255
         Left            =   4080
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Data 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Port 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Host 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Data to send:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Hostname:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   975
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
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3000
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub Connect_Click()
On Error Resume Next
Winsock1.LocalPort = Port.Text
Call Winsock1.Connect(Host.Text, Port.Text)
If Winsock1.State = 7 Then
Me.Caption = Me.Caption & " [" & Host.Text & "]"
Send.Enabled = True
End If
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 4230
Me.Width = 5400
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Terminate()
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Send_Click()
Winsock1.SendData Data.Text
CodeWin.Text = CodeWin.Text & Data.Text & vbCrLf
Data.Text = ""
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim NewData As String

Winsock1.GetData (NewData)

CodeWin.Text = CodeWin.Text & NewData & vbCrLf

End Sub
