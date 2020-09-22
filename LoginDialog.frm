VERSION 5.00
Begin VB.Form LoginDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1920
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Login Password :"
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
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   2775
   End
End
Attribute VB_Name = "LoginDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit  '  Form Login start
Private Sub CancelButton_Click()
Unload Me: End
End Sub
Private Sub Form_Activate()
Text1.SetFocus
End Sub
Private Sub OKButton_Click()
Dim lPass As String, rPass As String
lPass = Text1.Text
rPass = GetSetting("BasuDip", App.Title, "Login")
If Encrypt(lPass) = rPass Or rPass = vbNullString Then
Unload Me
LibProj.Show
Else
Text1.Text = vbNullString
Text1.SetFocus
End If
End Sub
Public Function Encrypt(ByVal strInput As String)
Dim iCount As Long, ingPtr As Long, strKey As String, CryptCode As String
strKey = StrReverse(strInput)
For iCount = 1 To Len(strInput)
CryptCode = CryptCode + Hex(Asc(Chr((Asc(Mid(strInput, iCount, 1))) Xor (Asc(Mid(strKey, ingPtr + 1, 1))))))
ingPtr = ((ingPtr + 1) Mod Len(strKey))
Next iCount
Encrypt = CryptCode
End Function
