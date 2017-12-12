VERSION 5.00
Begin VB.Form book_processing 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Book Processing"
   ClientHeight    =   8100
   ClientLeft      =   5655
   ClientTop       =   825
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame framebookinfo 
      Enabled         =   0   'False
      Height          =   5535
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   10335
      Begin VB.TextBox TxtBookId 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox TxtIsbn 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         TabIndex        =   1
         Top             =   240
         Width           =   3135
      End
      Begin VB.TextBox TxtAcc 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox TxtTitle 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   3
         Top             =   1680
         Width           =   4935
      End
      Begin VB.TextBox TxtAuthor 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2640
         TabIndex        =   4
         Top             =   2400
         Width           =   4935
      End
      Begin VB.ComboBox CmbClass 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "P.frx":0000
         Left            =   2640
         List            =   "P.frx":0013
         TabIndex        =   5
         Top             =   3120
         Width           =   4935
      End
      Begin VB.TextBox TxtBooksyn 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   2640
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ISBN :"
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
         Left            =   6000
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Accession No. :"
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
         Left            =   360
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Book ID No. :"
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
         Left            =   600
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Title Name :"
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
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Classification :"
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
         Left            =   360
         TabIndex        =   18
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Author's Name:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Book Synopsis :"
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
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   2055
      End
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
      Height          =   495
      Left            =   9000
      TabIndex        =   24
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
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
      Left            =   9000
      TabIndex        =   23
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "View/Print"
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
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   7200
      Width           =   1575
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
      Left            =   7320
      TabIndex        =   11
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "ADD"
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
      Left            =   2280
      TabIndex        =   8
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "DELETE"
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
      Left            =   5640
      TabIndex        =   10
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "SAVE"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "EDIT"
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
      TabIndex        =   9
      Top             =   7200
      Width           =   1335
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "UPDATE"
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
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "BOOK PROCESSING"
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
      Left            =   3720
      TabIndex        =   12
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "book_processing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsdata As ADODB.Recordset
Private Sub CmdAdd_Click()
framebookinfo.Enabled = True
TxtBookId.SetFocus
CmdAdd.Visible = False
CmdEdit.Enabled = False
CmdDel.Enabled = False
CmdExit.Enabled = False
CmdCancel.Visible = True
CmdExit.Visible = False
End Sub
Private Sub CmdCancel_Click()
Unload Me
book_processing.Show
End Sub
Private Sub CmdDel_Click()
Dim bookid As Integer
Dim del As Boolean
del = False
On Error GoTo l1
bookid = InputBox("Enter Book ID To Delete.")
Call connection.connection
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from bookprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
If bookid = rsdata(0) Then
TxtBookId.Text = rsdata(0)
TxtIsbn.Text = rsdata(1)
TxtAcc.Text = rsdata(2)
TxtTitle.Text = rsdata(3)
TxtAuthor.Text = rsdata(4)
CmbClass.Text = rsdata(5)
TxtBooksyn.Text = rsdata(6)
conf = MsgBox("Are you sure you want to delete this book's details?", vbYesNo)
    If conf = vbYes Then
    rsdata.Delete
    rsdata.Update
    rsdata.UpdateBatch
    rsdata.Close
    
    del = True
    Exit Do
    Else
    del = False
    TxtBookId.Text = ""
TxtIsbn.Text = ""
TxtAcc.Text = ""
TxtTitle.Text = ""
TxtAuthor.Text = ""
CmbClass.Text = ""
TxtBooksyn.Text = ""
TxtBookId.SetFocus
    Exit Do
    End If
Else
rsdata.MoveNext
End If

Loop

 If del = True Then
MsgBox "Book Deleted"
Unload Me
book_processing.Show
ElseIf del = False Then
MsgBox "Book ID Doesn't Exit."
End If
l1:
End Sub

Private Sub CmdEdit_Click()
Dim bookid As Integer
On Error GoTo l2
bookid = InputBox("Enter Book ID To Edit.")
Call connection.connection
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from bookprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
If bookid = rsdata(0) Then
TxtBookId.Text = rsdata(0)
TxtIsbn.Text = rsdata(1)
TxtAcc.Text = rsdata(2)
TxtTitle.Text = rsdata(3)
TxtAuthor.Text = rsdata(4)
CmbClass.Text = rsdata(5)
TxtBooksyn.Text = rsdata(6)
MsgBox "Now Edit Records And Update."
framebookinfo.Enabled = True
CmdEdit.Visible = False
CmdAdd.Enabled = False
CmdDel.Enabled = False
CmdExit.Enabled = False
CmdCancel.Visible = True
CmdExit.Visible = False
Exit Do

Else
rsdata.MoveNext
End If

Loop

If TxtBookId.Text = "" Then
MsgBox "No Book Found. Book ID Wrong."
End If
l2:
End Sub

Private Sub Cmdexit_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
Dim bookid As Integer
Dim found As Boolean
found = False
On Error GoTo l3
bookid = InputBox("Enter Book ID No. TO View Or Print.")
Call connection.connection
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from bookprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
If bookid = rsdata(0) Then
book_information.Sections("details").Controls("lblbid").Caption = rsdata(0)
book_information.Sections("details").Controls("Lblisbn").Caption = rsdata(1)
book_information.Sections("details").Controls("LblAccess").Caption = rsdata(2)
book_information.Sections("details").Controls("Lbltitle").Caption = rsdata(3)
book_information.Sections("details").Controls("lblauthor").Caption = rsdata(4)
book_information.Sections("details").Controls("Lblclassification").Caption = rsdata(5)
book_information.Sections("details").Controls("lblbooksynopsis").Caption = rsdata(6)
book_information.Show
   found = True
    Exit Do
       
Else
rsdata.MoveNext
End If
Loop
If found = False Then
MsgBox "Book ID Doesn't Exist."
End If
l3:
End Sub

Private Sub CmdReset_Click()
TxtBookId.Text = ""
TxtIsbn.Text = ""
TxtAcc.Text = ""
TxtTitle.Text = ""
TxtAuthor.Text = ""
CmbClass.Text = ""
TxtBooksyn.Text = ""
End Sub

Private Sub CmdSave_Click()
Call connection.connection

Set rsdata = New ADODB.Recordset
rsdata.Open "select *from bookprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
rsdata.MoveNext
Loop

rsdata.AddNew
rsdata(0) = TxtBookId.Text
rsdata(1) = TxtIsbn.Text
rsdata(2) = TxtAcc.Text
rsdata(3) = TxtTitle.Text
rsdata(4) = TxtAuthor.Text
rsdata(5) = CmbClass.Text
rsdata(6) = TxtBooksyn.Text

rsdata.Update
rsdata.UpdateBatch
rsdata.Close
conn.Close
MsgBox "Data Saved Successfully."
Unload Me
book_processing.Show
End Sub

Private Sub CmdUpdate_Click()
rsdata(0) = TxtBookId.Text
rsdata(1) = TxtIsbn.Text
rsdata(2) = TxtAcc.Text
rsdata(3) = TxtTitle.Text
rsdata(4) = TxtAuthor.Text
rsdata(5) = CmbClass.Text
rsdata(6) = TxtBooksyn.Text
rsdata.Update
rsdata.UpdateBatch

rsdata.Close
MsgBox "Record Updated Successfully."
Unload Me
book_processing.Show

End Sub

