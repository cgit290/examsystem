VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "KKBYCTC PRACTICAL EXAMINATION SYSTEM"
   ClientHeight    =   5190
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   7695
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6840
      Top             =   4440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu ext 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ext_Click()
'ans = InputBox("Enter the password to close?", "KKBYCTC")
'If (LCase(ans) = "chandan") Then
    End
'Else
 '   MsgBox "Wrong password", vbOKOnly
'End If
End Sub

Private Sub MDIForm_Load()
If Dir$("D:\EXAM", vbDirectory) = "" Then
MkDir ("D:\EXAM")
End If
loginfrm.Show
End Sub


