Attribute VB_Name = "Module1"
Public CN As New ADODB.Connection
Public rs As New ADODB.Recordset
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Sub main()
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\stdinformation.mdb;Persist Security Info=False"
CN.Open
MDIForm1.Show
End Sub
Sub refr(da As Object)
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open "SELECT * FROM sheet1", CN, adOpenDynamic, adLockOptimistic
Set da.DataSource = rs
End Sub


Public Function DisableCloseButton(frm As Form) As Boolean
Dim lHndSysMenu As Long
Dim lAns1 As Long, lAns2 As Long
lHndSysMenu = GetSystemMenu(frm.hWnd, 0)
lAns1 = RemoveMenu(lHndSysMenu, 6, MF_BYPOSITION)
lAns2 = RemoveMenu(lHndSysMenu, 5, MF_BYPOSITION)
DisableCloseButton = (lAns1 <> 0 And lAns2 <> 0)
End Function
