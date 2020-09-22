VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4755
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCommonFactor 
      Caption         =   "HCF"
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
      Left            =   2880
      TabIndex        =   30
      ToolTipText     =   "Common Factor"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton factorial 
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   29
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton correctEntry 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   28
      ToolTipText     =   "Clear Entry"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   27
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   26
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton back 
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "Symbol"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   25
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton memorycancel 
      Caption         =   "MC"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   24
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton memorycheck 
      Caption         =   "MR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   23
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton memoryminus 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   22
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton memoryplus 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   21
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cancel 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   20
      ToolTipText     =   "Cancel Operation"
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton reciprocal 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton divide 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton multiply 
      Caption         =   "×"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   17
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton plusminus 
      Caption         =   "±"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton dot 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   2040
      TabIndex        =   14
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   360
      TabIndex        =   13
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   12
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1200
      TabIndex        =   11
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton percent 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton exponen 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton minus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton plus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   360
      TabIndex        =   5
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2040
      TabIndex        =   4
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Number 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton equalto 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      ToolTipText     =   "Enter"
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label memory 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label display 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
   Begin VB.Menu mnucalculator 
      Caption         =   "&Calculator"
      Begin VB.Menu mnucopy 
         Caption         =   "C&opy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnucexit 
         Caption         =   "&Close"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuconstants 
      Caption         =   "Cons&tants"
      Begin VB.Menu mnupi 
         Caption         =   "&pi"
      End
      Begin VB.Menu musep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnue 
         Caption         =   "&e"
      End
   End
   Begin VB.Menu mnufunctions 
      Caption         =   "&Functions"
      Begin VB.Menu mnusin 
         Caption         =   "&sine (radian)"
      End
      Begin VB.Menu mnucos 
         Caption         =   "&cosine (radian)"
      End
      Begin VB.Menu mnutan 
         Caption         =   "&tangent (radian)"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnudtor 
         Caption         =   "degree --> &radian"
      End
      Begin VB.Menu mnurtod 
         Caption         =   "radian --> &degree"
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnulog10 
         Caption         =   "&logarithm base 10"
      End
      Begin VB.Menu mnuloge 
         Caption         =   "logarithm base &e"
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhindex 
         Caption         =   "&Index  . . ."
      End
      Begin VB.Menu mnusep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhabout 
         Caption         =   "&About  . . ."
      End
   End
   Begin VB.Menu mnupopup1 
      Caption         =   "popup1"
      Visible         =   0   'False
      Begin VB.Menu mnupopcopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnupoppaste 
         Caption         =   "Paste"
      End
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private operand1 As Double, operand2 As Double
Private operator As String, memoryVal As Double
Private clearDisplay As Boolean, appVer As String
Private Declare Function ShellAbout Lib "shell32.dll" Alias _
    "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, _
    ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Sub back_Click()
    If Len(display.Caption) > 0 Then display.Caption = Mid$(display.Caption, 1, Len(display.Caption) - 1)
End Sub
Private Sub cancel_Click()
    display.Caption = vbNullString
    operand1 = 0:    operand2 = 0
    operator = vbNullString
End Sub
Private Sub cmdCommonFactor_Click()
On Local Error Resume Next
Dim num1 As Long, num2 As Long, rHCF As Long
    num1 = InputBox("Input first No.", "HCF")
    num2 = InputBox("Input second No.", "HCF")
    If num1 <= 0 Or num2 <= 0 Then Exit Sub
    rHCF = HCFactor(num1, num2)
    MsgBox "Highest Common Factor for" & vbNewLine & num1 & _
        " and " & num2 & " is " & rHCF, vbInformation, "Result"
    display.Caption = rHCF
    clearDisplay = True
End Sub
Private Sub correctEntry_Click()
    display.Caption = vbNullString
End Sub
Private Sub display_Change()
    If Len(display.Caption) > 23 Then
        Call MsgBox("Data capacity is upto 23 digits" & vbNewLine & "Result may have error", , "Data OverFlow")
        Call back_Click
    End If
End Sub
Private Sub display_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnupopup1
End Sub
Private Sub divide_Click()
    operand1 = Val(display.Caption)
    operator = "/"
    display.Caption = vbNullString
End Sub
Private Sub dot_Click()
    If InStr(display.Caption, ".") Then
        Exit Sub
    Else
        display.Caption = display.Caption & "."
    End If
End Sub
Private Sub equalTo_Click()
Dim result As Double
    On Error GoTo eh:
    operand2 = Val(display.Caption)
    If operator = "+" Then result = operand1 + operand2
    If operator = "-" Then result = operand1 - operand2
    If operator = "*" Then result = operand1 * operand2
    If operator = "/" And operand2 <> 0 Then result = operand1 / operand2
    If operator = "^" Then result = operand1 ^ operand2
    display.Caption = result
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub exponen_Click()
    operand1 = Val(display.Caption)
    operator = "^"
    display.Caption = vbNullString
End Sub
Private Sub factorial_Click()
    Dim dispVal As Long
    dispVal = Val(display.Caption)
    If dispVal <= 170 And dispVal >= 0 Then
        display.Caption = FactNo(dispVal)
        clearDisplay = True
    Else
    MsgBox "Cannot compute factorial for " & dispVal
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case Chr$(KeyAscii)
    Case Is = "0":      Number_Click (0)
    Case Is = "1":      Number_Click (1)
    Case Is = "2":      Number_Click (2)
    Case Is = "3":      Number_Click (3)
    Case Is = "4":      Number_Click (4)
    Case Is = "5":      Number_Click (5)
    Case Is = "6":      Number_Click (6)
    Case Is = "7":      Number_Click (7)
    Case Is = "8":      Number_Click (8)
    Case Is = "9":      Number_Click (9)
    Case Is = "+":      plus_Click
    Case Is = "-":      minus_Click
    Case Is = "*":      multiply_Click
    Case Is = "/":      divide_Click
    Case Is = ".":      dot_Click
    Case Else
        If KeyAscii = vbKeyReturn Then
                        equalTo_Click
        ElseIf KeyAscii = vbKeyEscape Then
                        correctEntry_Click
        ElseIf KeyAscii = vbKeyBack Then
                        back_Click
        End If
    End Select
End Sub
Private Sub Form_Load()
    appVer = App.Major & "." & App.Minor
    frmCalculator.KeyPreview = True
End Sub
Private Sub memory_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    memory.ToolTipText = "Memory Value is : " & memoryVal
End Sub
Private Sub memorycancel_Click()
    memoryVal = 0
    clearDisplay = True
    Call CheckMemory
End Sub
Private Sub memorycheck_Click()
    display.Caption = memoryVal
    clearDisplay = True
    Call CheckMemory
End Sub
Private Sub memoryminus_Click()
    memoryVal = memoryVal - Val(display.Caption)
    clearDisplay = True
    Call CheckMemory
End Sub
Private Sub memoryplus_Click()
    memoryVal = memoryVal + Val(display.Caption)
    clearDisplay = True
    Call CheckMemory
End Sub
Private Sub minus_Click()
    operand1 = Val(display.Caption)
    operator = "-"
    display.Caption = vbNullString
End Sub
Private Sub mnucopy_Click()
    Clipboard.SetText (display.Caption)
End Sub
Private Sub mnucexit_Click()
    Unload Me
End Sub
Private Sub mnucos_Click()
    On Error GoTo eh:
    display.Caption = Cos(Val(display.Caption))
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnudtor_Click()
    On Error GoTo eh:
    display.Caption = Val(display.Caption) * 1.74532925199433E-02
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnue_Click()
    display.Caption = "2.71828182845904523536"
End Sub
Private Sub mnuhabout_Click()
Dim strAbout As String
    strAbout = "Calculator developed by Dipankar Basu" & vbNewLine & "http://www.geocities.com/basudip_in/"
    Call ShellAbout(Me.hwnd, App.Title, strAbout, Me.Icon)
End Sub
Private Sub mnuhindex_Click()
    SendKeys "{F1}", True
End Sub
Private Sub mnulog10_Click()
    On Error GoTo eh:
    If Val(display.Caption) > 0 Then display.Caption = Log(Val(display.Caption)) * 0.434294481903252
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnuloge_Click()
    On Error GoTo eh:
    If Val(display.Caption) > 0 Then display.Caption = Log(Val(display.Caption))
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnupi_Click()
    display.Caption = "3.141592653589793238463"
End Sub
Private Sub mnupopcopy_Click()
    Call mnucopy_Click
End Sub
Private Sub mnupoppaste_Click()
    If Clipboard.GetFormat(vbCFText) Then display.Caption = Val(Clipboard.GetText)
End Sub
Private Sub mnurtod_Click()
    On Error GoTo eh:
    display.Caption = Val(display.Caption) * 57.2957795130823
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnusin_Click()
    On Error GoTo eh:
    display.Caption = Sin(Val(display.Caption))
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub mnutan_Click()
    On Error GoTo eh:
    display.Caption = Tan(Val(display.Caption))
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub multiply_Click()
    operand1 = Val(display.Caption)
    operator = "*"
    display.Caption = vbNullString
End Sub
Private Sub Number_Click(Index As Integer)
    If clearDisplay Then
        display.Caption = vbNullString
        clearDisplay = False
    End If
    display.Caption = display.Caption + Number(Index).Caption
End Sub
Private Sub percent_Click()
Dim result As Double
    On Error GoTo eh:
    operand2 = Val(display.Caption)
    If operator = "+" Then result = operand1 + operand1 * operand2 / 100
    If operator = "-" Then result = operand1 - operand1 * operand2 / 100
    If operator = "*" Then result = operand1 * operand2 / 100
    If operator = "/" And operand2 <> 0 Then result = 100 / operand2
    If operator = "^" Then result = 0
    display.Caption = result
    clearDisplay = True
Exit Sub
eh:
    Call MsgBox(Err.Description, , Err.Source)
End Sub
Private Sub plus_Click()
    operand1 = Val(display.Caption)
    operator = "+"
    display.Caption = vbNullString
End Sub
Private Sub plusminus_Click()
    display.Caption = -Val(display.Caption)
End Sub
Private Sub reciprocal_Click()
    If Val(display.Caption) <> 0 Then
        display.ToolTipText = 1 / display.Caption
        display.Caption = 1 / Val(display.Caption)
    End If
End Sub
Private Sub CheckMemory()
    If memoryVal = 0 Then
        memory.Caption = vbNullString
        memorycancel.Enabled = False
    Else
        memory.Caption = "M"
        memorycancel.Enabled = True
    End If
End Sub
Private Function FactNo(ByVal myNumbr As Long) As Double
If myNumbr = 1 Or myNumbr = 0 Then
    FactNo = 1
Else
    FactNo = myNumbr * FactNo(myNumbr - 1)
End If
End Function
Private Function HCFactor(ByVal fNo As Long, ByVal sNo As Long) As Long
    If fNo < sNo Then HCFactor = HCFactor(sNo, fNo)
    If fNo Mod sNo = 0 Then
        HCFactor = sNo
    Else
        HCFactor = HCFactor(sNo, fNo Mod sNo)
    End If
End Function
