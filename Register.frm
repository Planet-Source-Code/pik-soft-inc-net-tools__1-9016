VERSION 5.00
Begin VB.Form Register 
   Caption         =   "Register Net Tools"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   Icon            =   "Register.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3840
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Registration Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3615
      Begin VB.CommandButton Cancel 
         Caption         =   "Cancel"
         Height          =   255
         Left            =   2280
         TabIndex        =   5
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Register 
         Caption         =   "Register"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Serial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1920
         Width           =   1815
      End
      Begin VB.TextBox ProCode 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox Email 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox NameTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Serial Number:"
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
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Product Code:"
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
         TabIndex        =   9
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Email:"
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
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
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
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cancel_Click()
Me.Hide
End Sub

Private Sub Form_Load()
ProCode.Text = ProductCode
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 3390
Me.Width = 3960
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Register_Click()
If Serial.Text = Serial_Check Then
Call SaveSetting("PiK Soft Net Tools", "Info", "Name", NameTxt.Text, HKEY_CURRENT_USER)
Call SaveSetting("PiK Soft Net Tools", "Info", "Email", Email.Text, HKEY_CURRENT_USER)
Call SaveSetting("PiK Soft Net Tools", "Info", "Product Code", ProCode.Text, HKEY_CURRENT_USER)
Call SaveSetting("PiK Soft Net Tools", "Info", "Regcode", Serial.Text, HKEY_CURRENT_USER)

Unload Me
End
End If
End Sub
