VERSION 5.00
Begin VB.Form frmStudent 
   Caption         =   "Student Register"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   10830
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
      Left            =   8160
      TabIndex        =   18
      ToolTipText     =   "Close"
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   5580
      TabIndex        =   17
      ToolTipText     =   "Ignore"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   4200
      TabIndex        =   16
      ToolTipText     =   "Update"
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
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
      Left            =   6960
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      Left            =   5580
      TabIndex        =   14
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      ToolTipText     =   "Move Last"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      ToolTipText     =   "Move Next"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Move Previous"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   9
      ToolTipText     =   "Move First"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   6720
      TabIndex        =   6
      Top             =   960
      Width           =   2295
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Height          =   495
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Go"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         MaxLength       =   5
         TabIndex        =   7
         ToolTipText     =   "Student ID"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2520
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1680
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Course Registered :"
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
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1740
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Student ID Number :"
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
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
End
Attribute VB_Name = "frmStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' Form new Student entry
Dim Cn As ADODB.Connection, Rs As ADODB.Recordset, _
Conn As String, QSQL As String
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdDelete_Click()
Dim response As Integer: On Error GoTo eh:
Display
response = MsgBox("Student Name: " & Rs!studentname & vbCrLf & _
    "Registered Course: " & Rs!Course, vbYesNo + vbDefaultButton2, "Delete Record")
If response = vbNo Then Exit Sub
With Rs
    .Delete
    .MovePrevious
    If .EOF Then .MoveLast
    If .BOF Then .MoveFirst
End With
Display
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End Sub
Private Sub cmdSearch_Click()
If Text4.Text = vbNullString Then Exit Sub
On Error GoTo eh:
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from LStudent where StudentID = " & _
      "'" & Text4.Text & "'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
MsgBox "Search Could not find any matching data", vbInformation, "Invalid Search Criteria"
GoTo CloseRsSearch:
End If
Text1.Text = RsSearch!Studentid
Text2.Text = RsSearch!studentname
Text3.Text = RsSearch!Course
CloseRsSearch:
RsSearch.Close: Set RsSearch = Nothing
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
Resume
End Sub
Private Sub Form_Activate()
On Error Resume Next
Display
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
    Rs.Open "lStudent", Cn, adOpenDynamic, adLockOptimistic, adCmdTable
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
 LibProj.mnuStudent.Enabled = False
 Unload Me
End Sub
Private Sub Form_Unload(cancel As Integer)
On Local Error Resume Next
LibProj.mnuTransaction.Enabled = True
Rs.Close: Set Rs = Nothing
Cn.Close: Set Cn = Nothing
End Sub
Private Sub Display()
On Error GoTo eh:
Text1.Text = Rs!Studentid
Text2.Text = Rs!studentname
Text3.Text = Rs!Course
Exit Sub
eh:
If Err.Number = 3021 Then
Call MsgBox("Data Base is empty, no current record" & vbCrLf & _
    "Click the ADD button to append a record data", , "Error")
Text1.Text = vbNullString
Text2.Text = vbNullString
Text3.Text = vbNullString
Else
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End If
End Sub
Private Sub cmdAdd_Click()
Rs.AddNew
Text1.Text = vbNullString: Text2.Text = vbNullString: Text3.Text = vbNullString
Text1.SetFocus
cmdAdd.Visible = False
cmdEdit.Visible = False
cmdSave.Visible = True
cmdCancel.Visible = True
cmdSearch.Enabled = False
cmdDelete.Enabled = False
End Sub
Private Sub CmdCancel_Click()
On Error Resume Next
Rs.CancelUpdate
cmdAdd.Visible = True
cmdEdit.Visible = True
cmdSave.Visible = False
cmdCancel.Visible = False
cmdSearch.Enabled = True
cmdDelete.Enabled = True
Display
End Sub
Private Sub cmdEdit_Click()
Display
Text1.SetFocus
cmdAdd.Visible = False
cmdEdit.Visible = False
cmdSave.Visible = True
cmdCancel.Visible = True
cmdSearch.Enabled = False
cmdDelete.Enabled = False
End Sub
Private Sub cmdFirst_Click()
Rs.MoveFirst: Display
End Sub
Private Sub cmdLast_Click()
Rs.MoveLast: Display
End Sub
Private Sub cmdNext_Click()
On Error Resume Next
With Rs
    .MoveNext
    If .EOF Then .MoveLast
End With
Display
End Sub
Private Sub cmdPrevious_Click()
On Error Resume Next
With Rs
    .MovePrevious
    If .BOF Then .MoveFirst
End With
Display
End Sub
Private Sub cmdSave_Click()
On Error GoTo eh:
Rs!Studentid = Text1.Text
Rs!studentname = Text2.Text
Rs!Course = Text3.Text
Rs.Update
cmdAdd.Visible = True
cmdEdit.Visible = True
cmdSave.Visible = False
cmdCancel.Visible = False
cmdSearch.Enabled = True
cmdDelete.Enabled = True
Display
Exit Sub
eh:
If Err.Number = -2147217900 Then
Call MsgBox("Duplicate entry exists, use a different Student ID Number", vbCritical, "Error")
Rs.CancelUpdate
Else
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End If
Resume Next
End Sub

