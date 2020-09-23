VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form SpeedChk 
   Caption         =   "PiK Soft Net Tools - Speed Check"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "SpeedChk.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   4335
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   120
         Top             =   120
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   1440
         Top             =   1200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label AverageS 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 KB/sec"
         Height          =   255
         Left            =   2040
         TabIndex        =   9
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Speedk 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 KB/sec"
         Height          =   255
         Left            =   2040
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Average Speed:"
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
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Speed:"
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
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Close 
         Caption         =   "Close"
         Height          =   255
         Left            =   3000
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton ChkSpeed 
         Caption         =   "Check Speed"
         Default         =   -1  'True
         Height          =   255
         Left            =   3000
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Host 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "SpeedChk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IP As String
Dim MS As Long, Sec As Integer
Dim Kbyte As String, Kbytes As Long

Dim iReturn As Long, sLowByte As String, sHighByte As String
Dim sMsg As String, HostLen As Long
Dim Hostent As Hostent, PointerToPointer As Long, ListAddress As Long
Dim WSAdata As WSAdata, DotA As Long, DotAddr As String, ListAddr As Long
Dim MaxUDP As Long, MaxSockets As Long, i As Integer
Dim Description As String, Status As String
' Ping Variables
Dim bReturn As Boolean, hIP As Long
Dim szBuffer As String
Dim Addr As Long
Dim RCode As String
Dim RespondingHost As String
' TRACERT Variables
Dim TraceRT As Boolean
Dim TTL As Integer
' WSock32 Constants
Const WS_VERSION_MAJOR = &H101 \ &H100 And &HFF&
Const WS_VERSION_MINOR = &H101 And &HFF&
Const MIN_SOCKETS_REQD = 0


Private Sub ChkSpeed_Click()
On Error Resume Next
If Winsock1.State <> sckClosed Then
    Winsock1.Close
End If
Kbyte = String$(10000, "0")
Sec = 0
MS = 0
Kbytes = 0
Timer1.Enabled = True
vbGetHostByName
Winsock1.Connect IP, 80
End Sub

Private Sub Close_Click()
Unload Me
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 3015
Me.Width = 4695
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Timer1_Timer()
MS = MS + 1
End Sub

Private Sub Winsock1_Connect()
Dim Times As Integer
Winsock1.SendData Kbyte
End Sub

Private Sub Winsock1_SendComplete()
Timer1.Enabled = False
MS = MS * 1000
Sec = MS / 1000
Kbytes = 100 / Sec
Kbytes = Len(Kbyte) / Sec
Kbytes = Kbytes / 1000
Speedk.Caption = Kbytes & " KB/sec"
AverageS.Caption = Kbytes & " KB/sec"
Winsock1.Close
End Sub
Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host.Text)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))
DoEvents
    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
DoEvents
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong6, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong6.Byte4)) + "." + CStr(Asc(IPLong6.Byte3)) _
        + "." + CStr(Asc(IPLong6.Byte2)) + "." + CStr(Asc(IPLong6.Byte1)))
    End If
End Sub


