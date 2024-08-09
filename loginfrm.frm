VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form loginfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6930
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   7095
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   600
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6600
      Top             =   6240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   6480
      Top             =   120
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "CLICK TO CONTINUE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4800
      Width           =   5655
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   600
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "kkb@001"
      Top             =   4080
      Width           =   5655
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Text            =   "yskkb07001@gmail.com"
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Show"
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   6000
      Width           =   5655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Department of Youth Services, Government of west Bengal"
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
      Left            =   720
      TabIndex        =   8
      Top             =   2520
      Width           =   5655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Online Practical Examination System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Training Centre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   2880
      Picture         =   "loginfrm.frx":0000
      Top             =   0
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kaktiya Bazar Youth"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1395
      TabIndex        =   5
      Top             =   1200
      Width           =   4245
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "loginfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rqst As String
Dim c As Integer
Private Sub Command1_Click()
If (Text1.Text <> "" And Text2.Text <> "") Then
    Dim lgData As String
    user_name = UCase(Text1.Text)
    Password = UCase(Text2.Text)
    lgData = user_name + "^" + Password
    If Winsock1.State <> sckConnected Then
        MsgBox "PLEASE RUNNING THE SERVER"
        Winsock1.Close
        Winsock1.Connect "DESKTOP-SI2P9P2", 2525
    ElseIf (user_name = "KKBYCTC" And Password = "86@CHANCHAL") Then
        marksfrm.Show
        Me.Hide
    Else
        Winsock1.SendData lgData
        Timer2.Enabled = True
        PBar.Visible = True
        Label7.Caption = "Validiting your login."
    End If
Else
    Label7.Caption = "Username and Password cannot be left blank!"
    Text1.SetFocus
End If
End Sub

Private Sub Form_Load()
c = 5
objcen Me
Winsock1.Connect "DESKTOP-SI2P9P2", 2525
Timer1.Enabled = True
data = Null
user_name = Null
Password = Null
sName = Null
sCourse = Null
sSem = Null
sExamEli = Null
sU1 = Null
sU2 = Null
sU3 = Null
PCODE = Null
ETIME = Null
DisableCloseButton MDIForm1
End Sub



Private Sub Label8_Click()
If (Label8.Caption = "Show") Then
    Label8.Caption = "Hide"
    Text2.PasswordChar = ""
Else
    Label8.Caption = "Show"
    Text2.PasswordChar = "*"
End If
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Winsock1.State <> sckConnected Then
    MDIForm1.Caption = "KKBYCTC Online Practical Examination System - Connection Status: Not Connected"
    Command1.Enabled = False
    Winsock1.Close
    Winsock1.Connect "DESKTOP-SI2P9P2", 2525
Else
    MDIForm1.Caption = "KKBYCTC Online Practical Examination System - Connection Status: Connected"
    Command1.Enabled = True
    Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
PBar.Value = PBar.Value + 5
If (PBar.Value > 50) Then
    If (rqst <> "NO") Then
        Label7.Caption = "Login verified. Logging in.."
        If PBar.Value = 95 Then
            PBar.Value = 0
            PBar.Visible = False
            Label7.Caption = ""
            Timer2.Enabled = False
            pcodefrm.Show
            loginfrm.Hide
        End If
    Else
        c = c - 1
        Label7.Caption = "The Login Id or Password is incorrect. " & c & " attempts left."
        PBar.Visible = False
        PBar.Value = 0
        Timer2.Enabled = False
        Text1.SetFocus
        If (c = 0) Then
            Command1.Enabled = False
            Label7.Caption = "Attempt failed. Please contact admin."
        End If
    End If
End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData data, vbString
rqst = data
If (data <> "NO") Then
        rData = Split(data, "^")
        sName = rData(0)
        sCourse = rData(1)
        sExamEli = rData(2)
        sU1 = rData(3)
        On Error Resume Next
        sU2 = rData(4)
        sU3 = rData(5)
        PCODE = rData(6)
        ETIME = rData(7)
        sSem = rData(8)
End If
End Sub

