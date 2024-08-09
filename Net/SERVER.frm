VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   10215
   ScaleWidth      =   18960
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid DBGrid1 
      Height          =   6495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11456
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7935
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton Command4 
         BackColor       =   &H008080FF&
         Caption         =   "Close"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   7320
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Settings"
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
         Left            =   0
         TabIndex        =   13
         Top             =   5880
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Refresh Data"
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
         Left            =   0
         TabIndex        =   12
         Top             =   6600
         Width           =   2415
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2580
         Left            =   0
         TabIndex        =   3
         Top             =   2520
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Double click for show password"
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0000FF00&
         Caption         =   "Start Server"
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
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5160
         Width           =   2415
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Index           =   0
         Left            =   360
         Top             =   4200
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Exam Duration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Admin Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Admin Username"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Connection List"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label lblpcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lbletime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblAdminUID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   1080
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim STDRECORD, USER_NAME, PW, sName, sCourse, sSem, sExamEli, sU1, sU2, sU3, sunit, PCODE, eTime, adminU, adminP As String
Dim GDATA() As String
Dim SCKNUMBER As Integer

Private Sub Command1_Click()
If (Command1.Caption = "Stop Server") Then
    For i = 0 To Winsock1.UBound Step 1
        Winsock1(i).Close
    Next
    Command1.Caption = "Start Server"
    Command1.BackColor = vbGreen
    MDIForm1.Caption = "KKBYCTC Server: Not Running"
    Command3.Enabled = True
    Command4.Enabled = True
    MDIForm1.ext.Enabled = True
    Label2.Caption = "Server-Not-Running"
    List1.Clear
Else
    Winsock1(0).Close
    Winsock1(0).LocalPort = 2525
    Winsock1(0).Listen
    Command1.Caption = "Stop Server"
    MDIForm1.Caption = "KKBYCTC Server: Running; Server-IP: " & Winsock1(0).LocalIP & "; Host-Name: " & Winsock1(0).LocalHostName
    Command3.Enabled = False
    Command4.Enabled = False
    MDIForm1.ext.Enabled = False
    Label2.Caption = "Server-Running"
    Command1.BackColor = vbRed
    List1.Clear
End If
adminU = lblAdminUID.Caption
adminP = Text1.Text
PCODE = Form3.Text2.Text
eTime = lbletime.Caption
End Sub

Private Sub Command2_Click()
refr DBGrid1
End Sub

Private Sub Command3_Click()
Form3.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
refr DBGrid1
Frame1.Left = Me.Width - Frame1.Width
DBGrid1.Width = Me.Width
DBGrid1.Height = Me.Height
Label2.Caption = Winsock1(0).LocalHostName
DisableCloseButton Me
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame1.Left = Me.Width - Frame1.Width - 500
DBGrid1.Width = Me.Width - 3000
DBGrid1.Height = Me.Height - 2000
End Sub

Private Sub Label4_Click()
If Text1.PasswordChar = "" Then
    Text1.PasswordChar = "*"
Else
    Text1.PasswordChar = ""
End If
End Sub

Private Sub Text1_Click()
Text1.PasswordChar = "*"
End Sub

Private Sub Text1_DblClick()
Text1.PasswordChar = ""
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Winsock1_ConnectionRequest(Index As Integer, ByVal requestID As Long)
SCKNUMBER = SCKNUMBER + 1
Load Winsock1(SCKNUMBER)
Winsock1(SCKNUMBER).Accept requestID
List1.AddItem Winsock1(Index).RemoteHostIP
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Set rs = New ADODB.Recordset
Winsock1(Index).GetData Data, vbString
GDATA = Split(Data, "^")
If (Len(Data) = 28) Then
        USER_NAME = GDATA(0)
        PW = GDATA(1)
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM sheet1 where username='" & USER_NAME & "'" & " AND PASSWORD='" & PW & "'", CN, adOpenDynamic, adLockOptimistic
        If (rs.EOF = True) Then
           Winsock1(Index).SendData "NO"
        Else
            sName = rs.Fields("NAME")
            sCourse = rs.Fields("COURSE")
            sExamEli = rs.Fields("EXAM_ELIGIBILITY")
            sU1 = rs.Fields("UNIT_1")
            sU2 = rs.Fields("UNIT_2")
            sU3 = rs.Fields("UNIT_3")
            'sSem = rs.Fields("SEMESTER")
            STDRECORD = sName + "^" + sCourse + "^" + sExamEli + "^" + sU1 + "^" + sU2 + "^" + sU3 + "^" + PCODE + "^" + eTime + "^" + sSem
            Winsock1(Index).SendData STDRECORD
        End If
ElseIf (Len(Data) = 23) Then
            rs.Open "select * from sheet1 where username='" & GDATA(0) & "'", CN, adOpenDynamic, adLockOptimistic
            sName = rs.Fields("name")
            sCourse = rs.Fields("course")
            sSem = rs.Fields("semester")
            If (rs.Fields("UNIT_1") <> "NO") Then
                If (rs.Fields("unit_3") = "YES") Then
                    sunit = "3"
                ElseIf (rs.Fields("UNIT_2") = "YES") Then
                    sunit = "2"
                ElseIf (rs.Fields("UNIT_1") = "YES") Then
                    sunit = "1"
                Else
                    sunit = "4"
                End If
                If (IsNumeric(GDATA(1)) = True) Then
                    rs.Fields("UNIT" & sunit & "_MARKS") = GDATA(1)
                    rs.Update
                Else
                    Winsock1(Index).SendData sName + "^" + sCourse + "^" + sSem + "^" + sunit
                End If
            Else
                sunit = "AB"
                 Winsock1(Index).SendData sName + "^" + sCourse + "^" + sSem + "^" + sunit
            End If
ElseIf (Len(Data) = 25) Then
        USER_NAME = GDATA(0)
        ENDTEST = GDATA(1)
        rs.Open "SELECT * FROM sheet1 where username='" & USER_NAME & "'", CN, adOpenDynamic, adLockOptimistic
        With rs
            If (ENDTEST = 1) Then
                .Fields("UNIT_1") = "YES"
                .Fields("DATE_U1") = Date
                .Fields("TIME_U1") = GDATA(2)
            ElseIf (ENDTEST = 2) Then
                .Fields("UNIT_2") = "YES"
                .Fields("DATE_U2") = Date
                .Fields("TIME_U2") = GDATA(2)
            ElseIf (ENDTEST = 3) Then
                .Fields("UNIT_3") = "YES"
                .Fields("DATE_U3") = Date
                .Fields("TIME_U3") = GDATA(2)
            End If
        End With
        rs.Update
Else
Winsock1(Index).SendData "NO"
End If
End Sub
