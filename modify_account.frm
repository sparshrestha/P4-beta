VERSION 5.00
Begin VB.Form modify_account 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Modify Account Name"
   ClientHeight    =   3930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   4080
      TabIndex        =   7
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox TxtPaswrd 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2040
      Width           =   3615
   End
   Begin VB.TextBox TxtAccName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1200
      Width           =   3615
   End
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
      Height          =   615
      Left            =   6240
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton CmdChange 
      Caption         =   "Change Account Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Enter Password To Confirm:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Type A New Account Name:"
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
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Modify Account Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "modify_account"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdChange_Click()
Call connection.connection
Dim rsdata As ADODB.Recordset
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from Authentication", conn, adOpenDynamic, adLockBatchOptimistic

If TxtPaswrd.Text = rsdata.Fields(1) And TxtAccName.Text <> "" Then


rsdata.Fields(0) = TxtAccName.Text
rsdata.UpdateBatch
MsgBox "Account Name Changed. New Account Name is: " & TxtAccName.Text
Unload Me

ElseIf TxtAccName = "" And Not TxtPaswrd = "" Then
MsgBox "New Account Name Is Not Entered."
ElseIf TxtPaswrd = "" And Not TxtAccName = "" Then
MsgBox "Current Password Is Not Entered."
ElseIf TxtAccName = "" And TxtPaswrd = "" Then
MsgBox "Both Fields Are Not Entered."
Else
MsgBox "Incorrect Password Entered."
End If
End Sub



Private Sub Cmdexit_Click()
Unload Me
End Sub

Private Sub CmdReset_Click()
TxtAccName.Text = ""
TxtPaswrd.Text = ""
TxtAccName.SetFocus
End Sub
