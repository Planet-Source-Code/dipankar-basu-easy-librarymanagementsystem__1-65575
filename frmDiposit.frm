VERSION 5.00
Begin VB.Form frmDiposit 
   Caption         =   "Diposit Book"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7920
   Begin VB.TextBox txtBookID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2340
      Width           =   3975
   End
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   9
      Top             =   1020
      Width           =   3975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdDiposit 
      Caption         =   "&Diposit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox txtBookID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   0
      ToolTipText     =   "Input Student Code"
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Book Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Top             =   2430
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Student's Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   1050
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Issue Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Book ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   1740
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Student's ID :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   1575
   End
End
Attribute VB_Name = "frmDiposit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' Form Diposit
Dim Cn As ADODB.Connection, Conn As String
Dim Rs As ADODB.Recordset
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDiposit_Click()
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from Ltrans where studID = " & _
    "'" & txtStudentID(0).Text & "' and BookStatus = 'I' and bookid = " & _
    "'" & txtBookID(0).Text & "'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic, adLockPessimistic
If RsSearch.EOF And RsSearch.BOF Then
Call MsgBox("Unable to retreive issued book details")
txtStudentID(0).SetFocus
Else
cmdDiposit.Enabled = True
txtDate.Text = FormatDateTime(RsSearch!tdate, vbLongDate)
Dim response As Integer
response = MsgBox("Is BookID : " & txtBookID(0).Text & _
    " submitted by StudentID : " & txtStudentID(0).Text _
    , vbQuestion + vbYesNo, "Submit Book Register")
If response = vbYes Then
RsSearch!BookStatus = "D"
RsSearch!tdate = FormatDateTime(Date, vbShortDate)
RsSearch.Update
With txtDate
    .Text = "BOOK IS DIPOSITED TO LIBRARY"
    .BorderStyle = 0
    .FontBold = True
    .ForeColor = vbGreen
End With
cmdDiposit.Enabled = False
cmdRefresh.SetFocus
End If
RsSearch.Close: Set RsSearch = Nothing
End If
End Sub
Private Sub cmdRefresh_Click()
Dim i As Integer
For i = 0 To 1 Step 1
txtStudentID(i).Text = vbNullString
txtBookID(i).Text = vbNullString
Next i
txtDate.Text = vbNullString
cmdDiposit.Enabled = False
txtStudentID(0).SetFocus
End Sub
Private Sub Form_Load()
On Error GoTo eh:
Conn = "Provider=SQLOLEDB.1;User ID=sa;Initial Catalog=LibDemo;Data Source=basudip"
       Set Cn = New ADODB.Connection
       With Cn
         .ConnectionString = Conn
         .CursorLocation = adUseClient
         .Open
       End With
    Set Rs = New ADODB.Recordset
    Rs.Open "lTrans", Cn, adOpenDynamic, adLockOptimistic, adCmdTable
Exit Sub
eh:
If Err.Number = -2147217865 Then
 MsgBox "Database Table does not exist", vbCritical, "Student Register Load Error"
 ElseIf Err.Number = -2147467259 Then
 MsgBox "SQL Server is not Started" & vbCrLf & _
    "check Control Panel Services Settings", vbCritical, "Please try after some time"
 Else
 MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
 End If
 LibProj.mnuTransaction.Enabled = False
 Unload Me
End Sub
Private Sub Form_Unload(cancel As Integer)
On Local Error Resume Next
Rs.Close: Set Rs = Nothing
Cn.Close: Set Cn = Nothing
End Sub
Private Sub txtStudentID_Change(Index As Integer)
If Index = 0 Then cmdDiposit.Enabled = False
End Sub
Private Sub txtStudentID_LostFocus(Index As Integer)
On Local Error Resume Next
If txtStudentID(0).Text = vbNullString Then Exit Sub
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select " & _
"lbook.bookid,lbook.bookname,lstudent.studentid,lstudent.studentname,ltrans.tdate from " & _
"lbook join ltrans   on (lbook.bookid=ltrans.bookid)" & _
"join lstudent on (lstudent.studentid=ltrans.studid)" & _
"where ltrans.bookstatus like '[i]' and " & _
"ltrans.studid = " & "'" & txtStudentID(0).Text & "'" & _
" order by lstudent.studentid,lbook.bookid"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
'Call MsgBox("Unable to retreive issued book details", , "Issue Register")
txtStudentID(1).Text = vbNullString
txtBookID(0).Text = vbNullString
txtBookID(1).Text = vbNullString
txtDate.Text = vbNullString
Else
cmdDiposit.Enabled = True
txtStudentID(1).Text = RsSearch!studentname
txtBookID(0).Text = RsSearch!bookid
txtBookID(1).Text = RsSearch!bookname
txtDate.Text = FormatDateTime(RsSearch!tdate, vbLongDate)
'CloseMe:
RsSearch.Close: Set RsSearch = Nothing
End If
End Sub
