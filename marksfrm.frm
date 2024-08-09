VERSION 5.00
Begin VB.Form marksfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   720
      Picture         =   "marksfrm.frx":0000
      ScaleHeight     =   3600
      ScaleWidth      =   8760
      TabIndex        =   14
      Top             =   4560
      Visible         =   0   'False
      Width           =   8760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   4560
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "YSKKB05"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   1
      Top             =   3720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "0 - 30"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9960
      TabIndex        =   15
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "@GMAIL.COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   13
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "*This box cannot be left blank."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "UNIT1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ADITA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CHANDAN BHANJA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "YSKKB06100@GMAIL.COM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   4920
      Picture         =   "marksfrm.frx":30B2
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "KAKTIYA BAZAR YOUTH COMPUTER TRAINING CENTRE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   600
      TabIndex        =   4
      Top             =   1680
      Width           =   9855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Practical Number Entry System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3240
      TabIndex        =   3
      Top             =   2160
      Width           =   4320
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Youth Services, Government of West Bengal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   1320
      Width           =   6720
   End
End
Attribute VB_Name = "marksfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim u1, u2, u3, ROLL, REG As String

Private Sub Command1_Click()
loginfrm.Winsock1.SendData REG + "^" + Text2.Text
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Label2.Caption = ""
Label3.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
End Sub

Private Sub Command3_Click()
Unload Me
loginfrm.Show
loginfrm.Text1.Text = ""
loginfrm.Text2.Text = ""
loginfrm.Text1.SetFocus

End Sub

Private Sub Form_Load()
Call objcen(Me)
End Sub



Private Sub Text1_LostFocus()
If (Text1.Text = "") Then
    Text1.SetFocus
Else
    REG = Text3.Text + Text1.Text + Label1.Caption
    loginfrm.Winsock1.SendData REG + "^" + "AA"
    Timer1.Enabled = True
    Picture1.Visible = True
End If
End Sub

Private Sub Text2_Change()
If (Val(Text2.Text) > 30 Or IsNumeric(Text2.Text) = False) Then
    Label10.ForeColor = vbred
    Text2.BackColor = vbred
    Command1.Enabled = False
Else
    Label10.ForeColor = vbBlack
    Text2.BackColor = vbWhite
    Label10.Caption = "0 - 30"
    Command1.Enabled = True
End If
End Sub



Private Sub Timer1_Timer()
If (Label2.Caption <> "") Then
    Picture1.Visible = False
    Timer1.Enabled = False
    If (sU1 <> "4") Then
        Label2.Caption = REG
        Label3.Caption = sName
        Label7.Caption = sCourse + "(" + sExamEli + ")"
        Label8.Caption = "UNIT" & sU1
    Else
        Label2.Caption = REG
        Label3.Caption = sName
        Label7.Caption = sCourse + "(" + sExamEli + ")"
        Label8.Caption = "AB"
        Text2.Text = "AB"
    End If
Else
    Picture1.Visible = True
    Timer1.Enabled = True
End If
End Sub
