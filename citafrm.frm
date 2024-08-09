VERSION 5.00
Begin VB.Form citafrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12855
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   Begin VB.VScrollBar VScroll1 
      Height          =   6495
      Left            =   12480
      TabIndex        =   13
      Top             =   2160
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   11160
      Picture         =   "citafrm.frx":0000
      ScaleHeight     =   2115
      ScaleWidth      =   2595
      TabIndex        =   12
      Top             =   9360
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12975
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   7320
         Top             =   600
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Time remaining                  Minutes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   3720
         TabIndex        =   11
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   5880
         TabIndex        =   10
         Top             =   480
         Width           =   135
      End
      Begin VB.Label scd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "59"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   270
         Left            =   6000
         TabIndex        =   9
         Top             =   480
         Width           =   270
      End
      Begin VB.Label mnt 
         BackStyle       =   0  'Transparent
         Caption         =   "29"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   5520
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image12 
         Height          =   1185
         Left            =   120
         Picture         =   "citafrm.frx":31B7
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CHANDAN BHANJA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   4
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "YSKKB05001@GMAIL.COM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   3
         Top             =   500
         Width           =   3375
      End
      Begin VB.Image Image11 
         Height          =   825
         Left            =   11880
         Picture         =   "citafrm.frx":9B75
         Top             =   120
         Width           =   825
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CITA(LEVEL1) - UNIT 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   2
         Top             =   800
         Width           =   3375
      End
   End
   Begin VB.CommandButton endt 
      BackColor       =   &H008080FF&
      Caption         =   "End Test"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   0
      Top             =   2160
      Width           =   12495
   End
End
Attribute VB_Name = "citafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Private Sub Command1_Click()
Call OPENAPP(Command1.Caption)
End Sub

Private Sub Command2_Click()
Call OPENAPP(Command2.Caption)
End Sub

Private Sub Command3_Click()
Call OPENAPP(Command3.Caption)
End Sub


Private Sub endt_Click()
Dim ans As String
ans = MsgBox("Are you sure to end the test", vbYesNo)
If (ans = vbYes) Then
    Call ENDTEST(Me, loginfrm.Winsock1)
End If
End Sub

Private Sub Form_Load()
objcen citafrm
Label9.Caption = sName
Label8.Caption = user_name
Label7.Caption = exminfofrm.Label3.Caption
a = 60
b = Val(ETIME) - 1
End Sub

Private Sub Timer1_Timer()
rTm = mnt.Caption
a = a - 1
If (a = 0) Then
     b = b - 1
    a = 60
ElseIf (a < 10) Then
    scd.Caption = "0" & a
Else
    scd.Caption = a
End If
If b = 0 Then
    Unload Me
    submsgfrm.Show
    Unload warningfrm
ElseIf (b < 10) Then
    mnt.Caption = "0" & b
Else
    mnt.Caption = b
End If
If (b = 10 And a = 1) Then
    MsgBox "Remaining Time: " & b & " Minutes", vbSystemModal + vbCritical, "Warning"
ElseIf (b = 5 And a = 1) Then
    MsgBox "Remaining Time: " & b & " Minutes", vbSystemModal + vbCritical, "Warning"
ElseIf (b = 1 And a = 1) Then
    MsgBox "Remaining Time: " & b & " Minutes", vbSystemModal + vbCritical, "Warning"
    Call ENDTEST(Me, loginfrm.Winsock1)
End If
End Sub

Private Sub VScroll1_Change()
'Image1.Top = -VScroll1.Value
End Sub
