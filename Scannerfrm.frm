VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ScannerFrm 
   Caption         =   "PiK Soft Net Tools - Port Scanner"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "Scannerfrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton optRemote 
         Caption         =   "Remote"
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   1560
         Width           =   855
      End
      Begin VB.OptionButton optLocal 
         Caption         =   "Local"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtIP 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtEndPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2880
         TabIndex        =   1
         Text            =   "65530"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtBeginPort 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Text            =   "1"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Default         =   -1  'True
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3000
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "To"
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
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Scan Port:"
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
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Current Port:"
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
         Left            =   2160
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblCurrent 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
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
         Left            =   3240
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   4335
      Begin VB.TextBox txtStatus 
         Height          =   2655
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   3000
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "ScannerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OnPort As Long
Dim LocalHost As Integer
Dim PortOpen As Long
Dim Host As String
Dim IP As String

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


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdStart_Click()
txtBeginPort.Enabled = False
txtEndPort.Enabled = False
cmdStart.Enabled = False
cmdStop.Enabled = True
txtStatus = ""
OnPort = txtBeginPort
PortDone = 0
cmdStop.SetFocus
Call Scanner(txtBeginPort, txtEndPort)
End Sub


Sub Scanner(Begin As Long, ending As Long)

TotalPorts = 0
PortOpen = 0
Do Until OnPort = txtEndPort
Pause 0.05
If PortDone = 1 Then lblCurrent = lblCurrent - 1: Exit Sub
DoEvents
lblCurrent = OnPort
If LocalHost = 1 Then
    If ScanPort(OnPort, Winsock1) = True Then
        TotalPorts = TotalPorts + 1
        PortOpen = PortOpen + 1
        If txtStatus = "" Then txtStatus = "Port " & OnPort & " is currently open.": GoTo thisPart
        txtStatus = txtStatus & vbCrLf & "Port " & OnPort & " is currently open."
        txtStatus.SelStart = Len(txtStatus)
    End If

ElseIf Len(txtIP.Text) > 1 Then
        Host = txtIP.Text
        vbGetHostByName
        Winsock1.Connect IP, OnPort
        Pause 0.2
        Winsock1.Close
End If

thisPart:
OnPort = OnPort + 1
Loop
lblCurrent = "Done"
txtStatus = txtStatus & vbCrLf & OnPort - 1 & " port(s) sucessfulley scanned." & vbCrLf & PortOpen & " Port(s) Open."
txtStatus.SelStart = Len(txtStatus)
cmdStop.Enabled = False
txtBeginPort.Enabled = True
txtEndPort.Enabled = True
cmdStart.Enabled = True
cmdStart.SetFocus
End Sub

Private Sub cmdStop_Click()
cmdStop.Enabled = False
txtBeginPort.Enabled = True
txtEndPort.Enabled = True
cmdStart.Enabled = True
PortDone = 1
txtStatus = txtStatus & vbCrLf & OnPort - 1 & " port(s) sucessfulley scanned." & vbCrLf & PortOpen & " Port(s) Open."
txtStatus.SelStart = Len(txtStatus)
cmdStart.SetFocus
End Sub

Private Sub Form_Load()
OnPort = 1
optLocal = True
LocalHost = 1
lblCurrent = "0"
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbNormal Then
Me.Height = 5670
Me.Width = 4695
ElseIf Me.WindowState = vbMaximized Then Me.WindowState = vbNormal
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Clean_Up
Winsock1.Close
End Sub

Private Sub optLocal_Click()
txtIP.Enabled = False
LocalHost = 1
End Sub

Private Sub optRemote_Click()
txtIP.Enabled = True
LocalHost = 2
End Sub

Private Sub Winsock1_Connect()
    txtStatus = txtStatus & vbCrLf & "Port " & OnPort & " is currently open."
    txtStatus.SelStart = Len(txtStatus)
    OnPort = OnPort + 1
    PortOpen = PortOpen + 1
End Sub
Public Sub vbGetHostByName()
    Dim szString As String
    Host = Trim$(Host)
    szString = String(64, &H0)
    Host = Host + Right$(szString, 64 - Len(Host))

    If gethostbyname(Host) = SOCKET_ERROR Then
        sMsg = "Winsock Error" & Str$(WSAGetLastError())
        MsgBox sMsg, 0, ""
    Else
        PointerToPointer = gethostbyname(Host) ' Get the pointer to the address of the winsock hostent structure
        CopyMemory Hostent.h_name, ByVal _
        PointerToPointer, Len(Hostent) ' Copy Winsock structure to the VisualBasic structure
        ListAddress = Hostent.h_addr_list ' Get the ListAddress of the Address List
        CopyMemory ListAddr, ByVal ListAddress, 4 ' Copy Winsock structure To the VisualBasic structure
        CopyMemory IPLong5, ByVal ListAddr, 4 ' Get the first list entry from the Address List
        CopyMemory Addr, ByVal ListAddr, 4
        IP = Trim$(CStr(Asc(IPLong5.Byte4)) + "." + CStr(Asc(IPLong5.Byte3)) _
        + "." + CStr(Asc(IPLong5.Byte2)) + "." + CStr(Asc(IPLong5.Byte1)))
    End If
End Sub


Sub Clean_Up()
On Error Resume Next
lblCurrent = 1
PortDone = 1
txtStatus = ""
cmdStop.Enabled = False
txtBeginPort.Enabled = True
txtEndPort.Enabled = True
cmdStart.Enabled = True
End Sub
