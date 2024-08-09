VERSION 5.00
Begin VB.Form insfrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "insfrm.frx":0000
      Top             =   0
      Width           =   10935
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10200
      Picture         =   "insfrm.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Agree above  terms and conditions"
      Enabled         =   0   'False
      Height          =   315
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Cilck on New Folder button"
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Please Save your question answer in 'D:\Exam\UserName' Folder. Click on Folder button to create folder."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   10335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on new folder button then  Check on agree"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "insfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Label2.Visible = False
End Sub

Private Sub Command1_Click()
If Check1.Value = 1 Then
    Unload Me
    exminfofrm.Show
Else
    Label2.Visible = True
End If
End Sub

Private Sub Command2_Click()
fName = user_name + "_" + sName
If Dir$("D:\EXAM\" & fName, vbDirectory) = "" Then
MkDir ("D:\EXAM\" & fName)
MsgBox "Folder creation Successful.", vbOKOnly
Check1.Enabled = True
Command2.Enabled = False
Else
    MsgBox fName & " Folder Already exsit.", vbOKOnly
    Check1.Enabled = True
    Command2.Enabled = False
End If
End Sub

Private Sub Form_Load()

Call objcen(Me)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
