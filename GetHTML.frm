VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form GetHTML 
   Caption         =   "PiK Soft Net Tools - Get HTML"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4815
   Icon            =   "GetHTML.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
      Begin RichTextLib.RichTextBox HTML 
         Height          =   1575
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   2778
         _Version        =   327680
         Enabled         =   -1  'True
         ScrollBars      =   2
         TextRTF         =   $"GetHTML.frx":0442
      End
      Begin InetCtlsObjects.Inet Inet1 
         Left            =   2520
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   327681
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton GetIt 
         Caption         =   "Get HTML"
         Default         =   -1  'True
         Height          =   255
         Left            =   3120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox URL 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "URL:"
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
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "GetHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 3600
Me.Width = 4935
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub GetIt_Click()
Dim Strsource As String
Strsource = Inet1.OpenURL(URL.Text)
HTML.Text = Strsource
End Sub

