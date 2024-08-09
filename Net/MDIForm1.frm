VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "KKBYCTC_Exam_Server"
   ClientHeight    =   6180
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11295
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu sett 
      Caption         =   "Set&tings"
   End
   Begin VB.Menu ext 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ext_Click()
End
End Sub

Private Sub MDIForm_Load()
Form1.Show
Form3.Show vbModal
DisableCloseButton Me
End Sub

Private Sub sett_Click()
Form3.Show
End Sub
