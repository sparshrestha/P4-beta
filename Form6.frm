VERSION 5.00
Begin VB.Form authentication 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "authentication"
   ClientHeight    =   4530
   ClientLeft      =   5370
   ClientTop       =   2340
   ClientWidth     =   7560
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      Picture         =   "Form6.frx":0000
      ScaleHeight     =   2025
      ScaleWidth      =   2025
      TabIndex        =   10
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      MaskColor       =   &H80000010&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton CmdReset 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.TextBox txtusername 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   0
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox txtpaswrd 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   25
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3000
      Width           =   5295
   End
   Begin VB.CommandButton CmdAccess 
      BackColor       =   &H0000FF00&
      Caption         =   "Access"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Library Management System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   240
      Width           =   4935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Authentication"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
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
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Password "
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
      Left            =   600
      TabIndex        =   6
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Input Your Username and Password to Authenticate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
End
Attribute VB_Name = "authentication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAccess_Click()
Call connection.connection
Dim rsdata As ADODB.Recordset
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from Authentication", conn, adOpenDynamic, adLockBatchOptimistic
If txtusername.Text = rsdata.Fields(0) And TxtPaswrd.Text = rsdata.Fields(1) Then

MsgBox "Access Granted. Welcome."
Unload Me
main_form.Show
ElseIf txtusername = "" And Not TxtPaswrd = "" Then
MsgBox "Username Is Not Entered."
txtusername.SetFocus
ElseIf TxtPaswrd = "" And Not txtusername = "" Then
MsgBox "Password Is Not Entered."
TxtPaswrd.SetFocus
ElseIf txtusername = "" And TxtPaswrd = "" Then
MsgBox "Both Fields Are Empty."
txtusername.SetFocus
Else
Call CmdReset_Click
MsgBox "ERROR. Incorrect Combination. Try Again."
txtusername.SetFocus
End If
End Sub

Private Sub Cmdexit_Click()
Unload Me
End Sub

Private Sub CmdReset_Click()
txtusername = ""
TxtPaswrd = ""
txtusername.SetFocus
End Sub

