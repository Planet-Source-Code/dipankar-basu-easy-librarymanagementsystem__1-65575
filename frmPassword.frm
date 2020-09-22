VERSION 5.00
Begin VB.Form frmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Login Password"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
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
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Close"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
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
      Left            =   2400
      TabIndex        =   6
      ToolTipText     =   "Reset Password"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Retype Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "New Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Password :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  ' Form Change Login Password
Private Sub Command1_Click()
On Error GoTo eh:
If Trim(Text1.Text) = vbNullString Then
Text1.Text = "Password"
ElseIf Len(Text2.Text) < 5 And Text2.Text <> vbNullString Then
Text2.SetFocus
Call MsgBox("Password should be atleast 5 chars in length", vbInformation, "Login Information")
Exit Sub
ElseIf Len(Text2.Text) > 25 Then
Call MsgBox("Password can be maximum 25 chars in length", vbInformation, "Login Information")
Exit Sub
End If
Dim rPass As String
rPass = GetSetting("BasuDip", App.Title, "Login")
If LoginDialog.Encrypt(Text1.Text) = rPass Or rPass = vbNullString Then
  If Text2.Text = Text3.Text Then
    If Text2.Text = vbNullString And rPass <> vbNullString Then
    Call DeleteSetting("BasuDip", App.Title, "Login")
    Call MsgBox("Password is successfully deleted", vbInformation, "Login Password Cleared")
    Unload Me
    Else
    rPass = LoginDialog.Encrypt(Text2.Text)
    Call SaveSetting("BasuDip", App.Title, "Login", rPass)
    Call MsgBox("Password is successfully changed", vbInformation, "Login Password Changed")
    Unload Me
    End If
  Else
  Call MsgBox("New Password do not match", vbInformation, "Password mismatch")
  End If
Else
Call MsgBox("Incorrect login Password", vbInformation, "Incorrect Login Information")
Text1.SetFocus
End If
Exit Sub
eh:
MsgBox Err.Source & " reports " & Err.Description, , "Error " & Err.Number
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Text1_GotFocus()
With Text1
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub Text2_GotFocus()
With Text2
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
Private Sub Text3_GotFocus()
With Text3
.SelStart = 0
.SelLength = Len(.Text)
End With
End Sub
