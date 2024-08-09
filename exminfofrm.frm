VERSION 5.00
Begin VB.Form exminfofrm 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   12900
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   12900
   ShowInTaskbar   =   0   'False
   Begin VB.Frame UNIT3FM 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   9000
      TabIndex        =   25
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080FF80&
         Caption         =   "Start Test"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   2640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   2640
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have already complete it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30 mins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   33
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   32
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   31
         Top             =   1200
         Width           =   495
      End
      Begin VB.Line Line7 
         X1              =   1560
         X2              =   1560
         Y1              =   960
         Y2              =   3120
      End
      Begin VB.Shape Shape3 
         Height          =   2175
         Left            =   240
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   30
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Taken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Assigned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "CITA - UNIT3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame UNIT2FM 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   5040
      TabIndex        =   15
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FF80&
         Caption         =   "Start Test"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "CITA - UNIT2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Assigned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Taken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   21
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Shape Shape2 
         Height          =   2175
         Left            =   240
         Top             =   960
         Width           =   2415
      End
      Begin VB.Line Line6 
         X1              =   1560
         X2              =   1560
         Y1              =   960
         Y2              =   3120
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   20
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30 mins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have already complete it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Line Line5 
         X1              =   240
         X2              =   2640
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line4 
         X1              =   240
         X2              =   2640
         Y1              =   2400
         Y2              =   2400
      End
   End
   Begin VB.Frame UNIT1FM 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   960
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FF80&
         Caption         =   "Start Test"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   2640
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   2640
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "You have already complete it."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "30 mins"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   11
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   1200
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   1560
         Y1              =   960
         Y2              =   3120
      End
      Begin VB.Shape Shape1 
         Height          =   2175
         Left            =   240
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Taken"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Test Assigned"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CITA - UNIT1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   6
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Sign out"
         Height          =   435
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label3 
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
         TabIndex        =   3
         Top             =   800
         Width           =   3375
      End
      Begin VB.Image Image2 
         Height          =   825
         Left            =   11880
         Picture         =   "exminfofrm.frx":0000
         Top             =   0
         Width           =   825
      End
      Begin VB.Label Label2 
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
         TabIndex        =   2
         Top             =   500
         Width           =   3375
      End
      Begin VB.Label Label1 
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
         TabIndex        =   1
         Top             =   120
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   1185
         Left            =   120
         Picture         =   "exminfofrm.frx":31B7
         Top             =   0
         Width           =   1185
      End
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"exminfofrm.frx":9B75
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
      Height          =   600
      Left            =   1080
      TabIndex        =   35
      Top             =   7200
      Visible         =   0   'False
      Width           =   10665
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a test to continue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2040
      Width           =   2895
   End
End
Attribute VB_Name = "exminfofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vbred As ColorConstants
Private Sub Command1_Click()
If sCourse = "CITA(LEVEL1)" Then
    citafrm.Command1.Caption = "NOTEPAD"
    citafrm.Command2.Caption = "MS PAINT"
    citafrm.Command3.Caption = "WORDPAD"
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\CITA_UNIT_1.jpg")
ElseIf sCourse = "DITA(LEVEL1)" Then
    citafrm.Command1.Caption = "MS_WORD"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\DITA_UNIT_1.jpg")
ElseIf (sCourse = "ADITA(LEVEL1)") Then
    citafrm.Command1.Caption = "NOTEPAD"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\ADITA_UNIT_1.jpg")
End If
Me.Hide
End Sub

Private Sub Command2_Click()
If (sCourse = "CITA") Then
    citafrm.Command1.Caption = "MS_WORD"
    citafrm.Command2.Caption = "MS_POWERPOINT"
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\CITA_UNIT_2.jpg")
ElseIf (sCourse = "DITA") Then
    citafrm.Command1.Caption = "MS_ACCESS"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\DITA_UNIT_2.jpg")
ElseIf (sCourse = "ADITA") Then
    citafrm.Command1.Caption = "NOTEPAD"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\ADITA_UNIT_2.jpg")
End If
Me.Hide
End Sub

Private Sub Command3_Click()
If (sCourse = "CITA") Then
    citafrm.Command1.Caption = "MS_EXCEL"
    citafrm.Command2.Caption = "VISUAL_FOXPRO"
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\CITA_UNIT_3.jpg")
ElseIf (sCourse = "DITA") Then
    citafrm.Command1.Caption = "VISUAL_BASIC"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\DITA_UNIT_3.jpg")
ElseIf (sCourse = "ADITA") Then
    citafrm.Command1.Caption = "VISUAL_C++"
    citafrm.Command2.Visible = False
    citafrm.Command3.Visible = False
    citafrm.Image1.Picture = LoadPicture(App.Path & "\Questions\ADITA_UNIT_3.jpg")
End If
Me.Hide
End Sub

Private Sub Command4_Click()
Unload Me
loginfrm.Show
loginfrm.Text1.SetFocus
End Sub

Private Sub Form_Load()
On Error Resume Next
pcodefrm.Text1.Text = ""
Label11.Caption = ETIME & "mins"
Label14.Caption = ETIME & "mins"
Label27.Caption = ETIME & "mins"
vbred = RGB(255, 100, 0)
Call objcen(exminfofrm)
Label1.Caption = sName
Label2.Caption = user_name
UNIT1FM.Visible = True
UC = 1
If (sU1 = "YES") Then
    Label5.Caption = Course & " - UNIT1"
    Command1.Caption = "Completed"
    Label12.Visible = True
    Label10.Caption = 1
    UNIT1FM.Enabled = False
    Command1.BackColor = vbred
    UNIT2FM.Visible = True
    UC = 2
    If (sU2 = "YES") Then
        Label20.Caption = Course & " - UNIT2"
        Command2.Caption = "Completed"
        Label13.Visible = True
        Label15.Caption = 1
        UNIT2FM.Enabled = False
        Command2.BackColor = vbred
        UNIT3FM.Visible = True
        UC = 3
        If (sU3 = "YES") Then
            Label20.Caption = Course & " - UNIT3"
            Command3.Caption = "Completed"
            Label28.Visible = True
            Label26.Caption = 1
            UNIT3FM.Enabled = False
            Command3.BackColor = vbred
        ElseIf (sExamEli = "NO") Then
            Command3.Caption = "Not Eligible"
            UNIT3FM.Enabled = False
            Command3.BackColor = vbYellow
            Label29.Visible = True
        End If
        ElseIf (sExamEli = "NO") Then
            Command2.Caption = "Not Eligible"
            UNIT2FM.Enabled = False
            Command2.BackColor = vbYellow
            Label29.Visible = True
    End If
    ElseIf (sExamEli = "NO") Then
            Command1.Caption = "Not Eligible"
            UNIT1FM.Enabled = False
            Command1.BackColor = vbYellow
            Label29.Visible = True
End If
Label3.Caption = sCourse & " - UNIT" & UC
Label5.Caption = sCourse & " - UNIT1"
Label20.Caption = sCourse & " - UNIT2"
Label21.Caption = sCourse & " - UNIT3"
End Sub

