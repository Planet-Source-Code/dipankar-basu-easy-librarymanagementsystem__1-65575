VERSION 5.00
Begin VB.Form frmNewBook 
   Caption         =   "New Book Entry"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8955
   Begin VB.CommandButton Command1 
      Caption         =   "C&lose"
      Height          =   495
      Left            =   7920
      TabIndex        =   16
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   15
      ToolTipText     =   "Ignore"
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   4440
      TabIndex        =   14
      ToolTipText     =   "Update"
      Top             =   2880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   495
      Left            =   5520
      TabIndex        =   12
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   2880
      Width           =   855
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
      Left            =   3480
      TabIndex        =   10
      ToolTipText     =   "Move Last"
      Top             =   2880
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
      Left            =   2640
      TabIndex        =   9
      ToolTipText     =   "Move Next"
      Top             =   2880
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
      Left            =   1800
      TabIndex        =   8
      ToolTipText     =   "Move Previous"
      Top             =   2880
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
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Move First"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   6600
      TabIndex        =   2
      Top             =   720
      Width           =   2175
      Begin VB.CommandButton cmdSearch 
         Caption         =   "Search"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Go"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
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
         Left            =   120
         MaxLength       =   5
         TabIndex        =   3
         ToolTipText     =   "Book ID Number"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.TextBox Text2 
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
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   1
      Top             =   1920
      Width           =   3615
   End
   Begin VB.TextBox Text1 
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
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Name"
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
      Left            =   840
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID Number"
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
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmNewBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' Form New Book Entry
Dim Cn As ADODB.Connection, Rs As ADODB.Recordset, _
Conn As String, QSQL As String
Private Sub cmdAdd_Click()
Rs.AddNew
Text1.Text = vbNullString: Text2.Text = vbNullString
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
Rs!bookid = Text1.Text
Rs!bookname = Text2.Text
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
Call MsgBox("Duplicate entry exists, use a different Book ID Number", vbCritical, "Error")
Rs.CancelUpdate
Else
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End If
Resume Next
End Sub
Private Sub cmdSearch_Click()
If Text3.Text = vbNullString Then Exit Sub
On Error GoTo eh:
Dim findStr As String, RsSearch As ADODB.Recordset
findStr = "select * from LBook where BookID = " & _
      "'" & Text3.Text & "'"
Set RsSearch = New ADODB.Recordset
RsSearch.Open findStr, Cn, adOpenDynamic
If RsSearch.EOF And RsSearch.BOF Then
MsgBox "Search Could not find any matching data", vbInformation, "Invalid Search Criteria"
GoTo CloseRsSearch:
End If
Text1.Text = RsSearch!bookid
Text2.Text = RsSearch!bookname
CloseRsSearch:
RsSearch.Close: Set RsSearch = Nothing
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
Resume
End Sub
Private Sub Command1_Click()
Unload Me
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
    Rs.Open "lBook", Cn, adOpenDynamic, adLockOptimistic, adCmdTable
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
 LibProj.mnuBook.Enabled = False
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
Text1.Text = Rs!bookid
Text2.Text = Rs!bookname
Exit Sub
eh:
If Err.Number = 3021 Then
Call MsgBox("Data Base is empty, no current record" & vbCrLf & _
    "Click the ADD button to append a record data", , "Error")
Text1.Text = vbNullString
Text2.Text = vbNullString
Else
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End If
End Sub
Private Sub cmdDelete_Click()
Dim response As Integer: On Error GoTo eh:
Display
response = MsgBox("Book ID: " & Rs!bookid & vbCrLf & _
    "Book Name: " & Rs!bookname, vbYesNo + vbDefaultButton2, "Delete Record")
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
