VERSION 5.00
Begin VB.Form Serial 
   Caption         =   "PiK Soft Net Tools - Registration Maker"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   Icon            =   "Serial.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Exit 
         Caption         =   "Exit"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Generate 
         Caption         =   "Generate"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Serial 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Procode 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Email 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox NameTxt 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Serial:"
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
         Top             =   1800
         Width           =   735
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
         Left            =   240
         TabIndex        =   5
         Top             =   1320
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
         Left            =   240
         TabIndex        =   3
         Top             =   840
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Serial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Exit_Click()
End
End Sub

Function Serial_Check() As String
Dim I As Integer
Dim Letter As String, code As String, Ser As Long, Sertxt As String

If Len(NameTxt.Text) < Len(Email.Text) Then

For I = 1 To Len(NameTxt.Text)

    Letter = Asc(Mid(NameTxt.Text, I, 1))
    code = Asc(Mid(Email.Text, I, 1))
    Letter = Letter Mod code
    Sertxt = Procode.Text * (Asc(Letter) / 1.3)
Next I

ElseIf Len(NameTxt.Text) = Len(Email.Text) Then

For I = 1 To Len(NameTxt.Text)

    Letter = Asc(Mid(NameTxt.Text, I, 1))
    code = Asc(Mid(Email.Text, I, 1))
    Letter = Letter Mod code
    Sertxt = Procode.Text * (Asc(Letter) / 1.3)
Next I

ElseIf Len(NameTxt.Text) > Len(Email.Text) Then

For I = 1 To Len(Email.Text)

    Letter = Asc(Mid(NameTxt.Text, I, 1))
    code = Asc(Mid(Email.Text, I, 1))
    Letter = Letter Mod code
    Sertxt = Procode.Text * (Asc(Letter) / 1.3)
Next I
End If
Sertxt = ReplaceString(Sertxt, ".", "")
Sertxt = ReplaceString(Sertxt, "+", "")
Serial_Check = Sertxt
End Function

Private Sub Generate_Click()
Serial.Text = Serial_Check
End Sub

Function ReplaceString(MyString As String, ToFind As String, ReplaceWith As String) As String
  Dim Spot As Long, NewSpot As Long, LeftString As String
    Dim RightString As String, NewString As String
    Spot& = InStr(LCase(MyString$), LCase(ToFind))
    NewSpot& = Spot&
    Do
        If NewSpot& > 0& Then
            LeftString$ = Left(MyString$, NewSpot& - 1)
            If Spot& + Len(ToFind$) <= Len(MyString$) Then
                RightString$ = Right(MyString$, Len(MyString$) - NewSpot& - Len(ToFind$) + 1)
            Else
                RightString = ""
            End If
            NewString$ = LeftString$ & ReplaceWith$ & RightString$
            MyString$ = NewString$
        Else
            NewString$ = MyString$
        End If
        Spot& = NewSpot& + Len(ReplaceWith$)
        If Spot& > 0 Then
            NewSpot& = InStr(Spot&, LCase(MyString$), LCase(ToFind$))
        End If
    Loop Until NewSpot& < 1
    ReplaceString$ = NewString$
End Function
