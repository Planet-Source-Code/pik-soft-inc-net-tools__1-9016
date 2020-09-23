VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form WhoisFrm 
   Caption         =   "PiK Soft Net Tools - Whois"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   Icon            =   "WhoisFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4695
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox ServerTxt 
      Height          =   315
      Left            =   1080
      TabIndex        =   0
      Text            =   "198.41.0.6"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   4455
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   2880
         Top             =   960
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.TextBox txtWhois 
         Height          =   1335
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3360
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdWhois 
         Caption         =   "Whois"
         Default         =   -1  'True
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Server:"
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
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Domain:"
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
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
End
Attribute VB_Name = "WhoisFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Close_Click()
Unload Me
End Sub

Private Sub cmdWhois_Click()
Winsock1.Close
Dim WhoisStr As String
txtWhois.Text = ""
Winsock1.Connect Servertxt, 43
End Sub

Private Sub Form_Load()
Servertxt.AddItem "198.41.0.8"
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 3600
Me.Width = 4815
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_Connect()
On Error Resume Next
Winsock1.SendData ("whois " & Host.Text & vbCrLf)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dataA
Winsock1.GetData dataA, vbString
txtWhois.Text = txtWhois.Text & dataA '& vbCrLf
Dim counter As Long
counter = 1
start:
   Dim Search, where   ' Declare variables.
   ' Get search string from user.
   Search = Chr$(10)
   where = InStr(counter, txtWhois.Text, Search, vbTextCompare) ' Find string in text.
   'MsgBox Where
   If where Then   ' If found,
      txtWhois.SelStart = where - 1   ' set selection start and
      txtWhois.SelLength = Len(Search)
      txtWhois.SelText = vbCrLf
      counter = where + txtWhois.SelLength + 2 ': 'MsgBox counter
   Else
      Exit Sub  ' Notify user.
   End If

GoTo start
Winsock1.Close
End Sub





