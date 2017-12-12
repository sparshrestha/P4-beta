VERSION 5.00
Begin VB.MDIForm main_form 
   BackColor       =   &H8000000C&
   Caption         =   $"main_form.frx":0000
   ClientHeight    =   10080
   ClientLeft      =   225
   ClientTop       =   255
   ClientWidth     =   14715
   LinkTopic       =   "MDIForm1"
   Picture         =   "main_form.frx":0098
   WindowState     =   2  'Maximized
   Begin VB.Menu bookinfo 
      Caption         =   "Book Processing"
   End
   Begin VB.Menu bookissue 
      Caption         =   "Book Issue"
   End
   Begin VB.Menu bookreturn 
      Caption         =   "Book Return"
   End
   Begin VB.Menu minfo 
      Caption         =   "Membership Processing"
   End
   Begin VB.Menu modify_accountname 
      Caption         =   "Modify Account"
   End
   Begin VB.Menu changepass 
      Caption         =   "Modify Password"
   End
   Begin VB.Menu unload_me 
      Caption         =   "Close"
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bookinfo_Click()
book_processing.Show

End Sub

Private Sub bookissue_Click()
book_issue.Show

End Sub

Private Sub bookreturn_Click()
book_return.Show

End Sub

Private Sub changepass_Click()
modify_password.Show

End Sub


Private Sub minfo_Click()
membership_processing.Show

End Sub

Private Sub modify_accountname_Click()
modify_account.Show

End Sub

Private Sub unload_me_Click()
Unload Me
End Sub
