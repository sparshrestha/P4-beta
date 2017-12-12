VERSION 5.00
Begin VB.Form modify_password 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Password"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton CmdChangepaswrd 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   3480
      Width           =   2295
   End
   Begin VB.TextBox TxtConfnewpaswrd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox TxtInnewpaswrd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox TxtInoldpaswrd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3360
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Confirm New Password :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Input New Password :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Input Old Password :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Modify Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "modify_password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdChangepaswrd_Click()
Call connection.connection
Dim rsdata As ADODB.Recordset
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from Authentication", conn, adOpenDynamic, adLockBatchOptimistic

If TxtInoldpaswrd.Text = rsdata.Fields(1) And TxtInnewpaswrd.Text = TxtConfnewpaswrd.Text Then


rsdata.Fields(1) = TxtInnewpaswrd.Text
rsdata.UpdateBatch
MsgBox "Password Changed Successfully."
Unload Me

ElseIf TxtInoldpaswrd = "" And Not TxtInnewpaswrd.Text = "" And Not TxtConfnewpaswrd.Text = "" Then
MsgBox "Old Password Is Not Entered."
TxtInoldpaswrd.SetFocus
ElseIf TxtInnewpaswrd.Text = "" And Not TxtInoldpaswrd.Text = "" And Not TxtConfnewpaswrd.Text = "" Then
MsgBox "New Password Is Not Entered."
TxtInnewpaswrd.SetFocus
ElseIf TxtConfnewpaswrd.Text = "" And Not TxtInoldpaswrd.Text = "" And Not TxtInnewpaswrd.Text = "" Then
MsgBox "Confirm Password."
TxtConfnewpaswrd.SetFocus
ElseIf TxtInoldpaswrd.Text = "" And TxtInnewpaswrd.Text = "" And Not TxtConfnewpaswrd.Text = "" Then
MsgBox "Old Password And New Password Are Not Entered."
TxtInoldpaswrd.SetFocus
ElseIf TxtInoldpaswrd.Text = "" And TxtConfnewpaswrd.Text = "" And Not TxtInnewpaswrd = "" Then
MsgBox "Old Password And Confirmation Password Are Not Entered."
TxtInoldpaswrd.SetFocus
ElseIf TxtInnewpaswrd.Text = "" And TxtConfnewpaswrd.Text = "" And Not TxtInoldpaswrd.Text = "" Then
MsgBox "New Password And Confirmation Password Are Not Entered."
TxtInnewpaswrd.SetFocus
ElseIf TxtInoldpaswrd.Text = "" And TxtInnewpaswrd.Text = "" And TxtConfnewpaswrd.Text = "" Then
MsgBox "No Data Entered."
TxtInoldpaswrd.SetFocus
Else
MsgBox "Wrong Inputs."
Call CmdReset_Click
TxtInoldpaswrd.SetFocus
End If
End Sub

Private Sub Cmdexit_Click()
Unload Me

End Sub

Private Sub CmdReset_Click()
TxtInnewpaswrd = ""
TxtInoldpaswrd = ""
TxtConfnewpaswrd = ""
TxtInoldpaswrd.SetFocus
End Sub

