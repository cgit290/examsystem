Attribute VB_Name = "Module1"
Global data, user_name, Password, sName, sCourse, sSem, sExamEli, sU1, sU2, sU3, OSBIT, PCODE, ETIME, UC As String
Global rTm As Integer
Global rData() As String

Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&

Public Function DisableCloseButton(frm As Form) As Boolean
Dim lHndSysMenu As Long
Dim lAns1 As Long, lAns2 As Long
lHndSysMenu = GetSystemMenu(frm.hWnd, 0)
lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function

Function objcen(obj As Object)
obj.Top = Screen.Height / 2 - obj.Height / 2
obj.Left = Screen.Width / 2 - obj.Width / 2
End Function

Function OPENAPP(NAME As String)
OSBIT = "C:\Program Files"
On Error Resume Next
If NAME = "NOTEPAD" Then
    a = Shell("C:\Windows\system32\NOTEPAD.exe", vbMaximizedFocus)
ElseIf (NAME = "MS PAINT") Then
    a = Shell("C:\Windows\system32\MSPAINT.exe", vbMaximizedFocus)
ElseIf (NAME = "WORDPAD") Then
    a = Shell("C:\Windows\system32\WRITE.exe", vbMaximizedFocus)
ElseIf (NAME = "MS_EXCEL") Then
    a = Shell(OSBIT & "\Microsoft Office\Office15\excel.EXE", vbMaximizedFocus)
ElseIf (NAME = "MS_POWERPOINT") Then
    a = Shell(OSBIT & "\Microsoft Office\Office15\powerpnt.EXE", vbMaximizedFocus)
ElseIf (NAME = "MS_WORD") Then
    a = Shell(OSBIT & "\Microsoft Office\Office15\WINWORD.EXE", vbMaximizedFocus)
ElseIf (NAME = "VISUAL_FOXPRO") Then
    a = Shell(OSBIT & "\Microsoft Visual Studio\Vfp98\VFP6.EXE", vbMaximizedFocus)
ElseIf (NAME = "VISUAL_BASIC") Then
    a = Shell(OSBIT & "\Microsoft Visual Studio\VB98\VB6.EXE", vbMaximizedFocus)
ElseIf (NAME = "MS_ACCESS") Then
    a = Shell(OSBIT & "\Microsoft Office\Office15\MSACCESS.EXE", vbMaximizedFocus)
ElseIf (NAME = "MS_ACCESS") Then
    a = Shell(OSBIT & "\Microsoft Visual Studio\Common\MSDev98\Bin\MSDEV.EXE", vbNormalFocus)
End If
End Function

Function ENDTEST(objfrm As Object, wnck As Object)
    Unload objfrm
    submsgfrm.Show
    If (UC = 1) Then
        wnck.SendData user_name & "^" & UC & "^" & rTm
    ElseIf (UC = 2) Then
        wnck.SendData user_name & "^" & UC & "^" & rTm
    ElseIf (UC = 3) Then
        wnck.SendData user_name & "^" & UC & "^" & rTm
    End If
End Function


