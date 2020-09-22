VERSION 5.00
Begin VB.MDIForm LibProj 
   BackColor       =   &H8000000C&
   Caption         =   "My Library management System"
   ClientHeight    =   3660
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6510
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuRegister 
      Caption         =   "Register"
      Begin VB.Menu mnuBook 
         Caption         =   "Book"
      End
      Begin VB.Menu mnuStudent 
         Caption         =   "Student"
      End
      Begin VB.Menu mnuSepR1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "Transaction"
      Begin VB.Menu mnuIssue 
         Caption         =   "Issue"
      End
      Begin VB.Menu mnuDiposit 
         Caption         =   "Diposit"
      End
      Begin VB.Menu mnuSepT1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDueList 
         Caption         =   "View &Due List"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Change Login &Password"
      End
      Begin VB.Menu mnuSepH1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalculator 
         Caption         =   "Calculator"
      End
      Begin VB.Menu mnuSepH2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About . . ."
      End
   End
End
Attribute VB_Name = "LibProj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub mnuAbout_Click()
Dim strMsg As String
strMsg = "Library Management System by Dipankar Basu" & vbCrLf & _
        "http://www.geocities.com/basudip_in/"
MsgBox strMsg, vbOKOnly, "Demo Library"
End Sub
Private Sub mnuBook_Click()
On Local Error Resume Next
frmNewBook.Show
LibProj.mnuTransaction.Enabled = False
End Sub
Private Sub mnuCalculator_Click()
frmCalculator.Show
End Sub
Private Sub mnuContents_Click()
On Error Resume Next
App.HelpFile = App.Path & "\" & "libproj.chm"
SendKeys "{F1}", True
End Sub
Private Sub mnuDiposit_Click()
On Local Error Resume Next
frmDiposit.Show
End Sub
Private Sub mnuExit_Click()
Unload Me: End
End Sub
Private Sub mnuIssue_Click()
On Local Error Resume Next
frmIssue.Show
End Sub
Private Sub mnuPassword_Click()
frmPassword.Show vbModal, Me
End Sub
Private Sub mnuStudent_Click()
On Local Error Resume Next
frmStudent.Show
LibProj.mnuTransaction.Enabled = False
End Sub
