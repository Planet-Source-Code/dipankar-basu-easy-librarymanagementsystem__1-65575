VERSION 5.00
Begin VB.Form frmIssue 
   Caption         =   "Issue Book"
   ClientHeight    =   6435
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7860
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7860
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "Issue &Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3540
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Top             =   4080
      Width           =   4335
   End
   Begin VB.TextBox txtBookName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3240
      Width           =   4335
   End
   Begin VB.TextBox txtBookID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2400
      Width           =   4335
   End
   Begin VB.TextBox txtStudentName 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox txtStudentID 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label5 
      Caption         =   "Issue Date :"
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
      Left            =   480
      TabIndex        =   9
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Book Name :"
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
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Book ID Number :"
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
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Student's Name :"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID Number :"
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
      Left            =   480
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit   ' Form Issue
Dim Cn As ADODB.Connection, Rs As ADODB.Recordset, _
Conn As String, QSQL As String
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdIssue_Click()
If txtStudentName.Text = vbNullString Or _
    txtBookName.Text = vbNullString Or _
    txtDate.Text = vbNullString Then Exit Sub
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from Ltrans where studID = " & _
    "'" & txtStudentID.Text & "'" & " and BookStatus = 'I'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
Call IssueBook
Else
MsgBox "A Book is issued to Student ID :  " & RsSearch!studid & vbCrLf & _
    "Issued Book ID is :  " & RsSearch!bookid & vbCrLf & _
    "Issue Date is :  " & Format(RsSearch!tdate, "MMMM dd,yyyy."), vbInformation, "Book Issued Register"
End If
RsSearch.Close: Set RsSearch = Nothing
End Sub
Private Sub cmdRefresh_Click()
txtStudentID.Text = vbNullString
txtStudentName.Text = vbNullString
txtBookID.Text = vbNullString
txtBookName.Text = vbNullString
txtDate.Text = vbNullString
cmdIssue.Enabled = True
txtStudentID.SetFocus
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
Private Sub txtBookID_LostFocus()
If txtBookID.Text = vbNullString Then
cmdIssue.Enabled = False
Exit Sub
Else
cmdIssue.Enabled = True
End If
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from LBook where BookID = " & _
      "'" & txtBookID.Text & "'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
MsgBox "Search Could not find any matching data", vbInformation, "Invalid Book ID"
txtBookID.Text = vbNullString
cmdIssue.Enabled = False
GoTo CloseRsSearch:
End If
txtBookName.Text = RsSearch!bookname
CloseRsSearch:
RsSearch.Close: Set RsSearch = Nothing
Exit Sub
End Sub
Private Sub txtDate_GotFocus()
If IsDate(txtDate.Text) = False Then txtDate.Text = Format(Date, "MMMM dd, yyyy")
End Sub
Private Sub txtStudentID_LostFocus()
If txtStudentID.Text = vbNullString Then
cmdIssue.Enabled = False
Exit Sub
Else
cmdIssue.Enabled = True
End If
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from LStudent where StudentID = " & _
      "'" & txtStudentID.Text & "'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
MsgBox "Search Could not find any matching data", vbInformation, "Invalid Student ID"
txtStudentID.Text = vbNullString
cmdIssue.Enabled = False
GoTo CloseRsSearch:
End If
txtStudentName.Text = RsSearch!studentname
CloseRsSearch:
RsSearch.Close: Set RsSearch = Nothing
Exit Sub
End Sub
Private Sub IssueBook()
On Error GoTo eh:
Rs.AddNew
Rs!bookid = txtBookID.Text
Rs!studid = txtStudentID.Text
Rs!tdate = CStr(Format(txtDate.Text, "dd/mm/yy"))
Rs!BookStatus = "I"
Rs.Update
MsgBox Trim(txtBookName.Text) & " is issued to " & Trim(txtStudentName.Text) & " on " & Format(Rs!tdate, "MMMM dd, yyyy."), vbInformation, "Book is Issued"
Call cmdRefresh_Click
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
Rs.CancelUpdate
End Sub
