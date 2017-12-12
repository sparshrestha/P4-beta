VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form membership_processing 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Membership Processing"
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   13710
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
      Left            =   11400
      TabIndex        =   37
      Top             =   6840
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
      Left            =   1200
      TabIndex        =   36
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame framememinfo 
      Enabled         =   0   'False
      Height          =   5415
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   13575
      Begin VB.OptionButton OptMale 
         Caption         =   "Male"
         Height          =   495
         Left            =   9840
         TabIndex        =   35
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton OptFemale 
         Caption         =   "Female"
         Height          =   495
         Left            =   11400
         TabIndex        =   34
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox TxtMbId 
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
         Left            =   2520
         TabIndex        =   20
         Top             =   120
         Width           =   3495
      End
      Begin VB.TextBox TxtFname 
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
         Left            =   2520
         TabIndex        =   19
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtMname 
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
         Left            =   6720
         TabIndex        =   18
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtLname 
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
         Left            =   11160
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox TxtOccu 
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
         Left            =   2520
         TabIndex        =   16
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox TxtEmail 
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
         Left            =   2520
         TabIndex        =   15
         Top             =   2280
         Width           =   4575
      End
      Begin VB.TextBox TxtStreet 
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
         Left            =   2520
         TabIndex        =   14
         Top             =   3000
         Width           =   3495
      End
      Begin VB.TextBox TxtCity 
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
         Left            =   2520
         TabIndex        =   13
         Top             =   3720
         Width           =   3495
      End
      Begin VB.TextBox TxtState 
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
         Left            =   2520
         TabIndex        =   12
         Top             =   4440
         Width           =   3495
      End
      Begin VB.TextBox TxtCell 
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
         Left            =   9840
         TabIndex        =   11
         Top             =   2280
         Width           =   3255
      End
      Begin VB.TextBox TxtHome 
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
         Left            =   9840
         TabIndex        =   10
         Top             =   3000
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPDOB 
         Height          =   495
         Left            =   9840
         TabIndex        =   9
         Top             =   3720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Format          =   102563841
         CurrentDate     =   41171
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Member ID No:"
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
         TabIndex        =   33
         Top             =   120
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "First Name:"
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
         TabIndex        =   32
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Gender :"
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
         Left            =   8160
         TabIndex        =   31
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Street:"
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
         Left            =   1080
         TabIndex        =   30
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Cell phone :"
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
         Left            =   7800
         TabIndex        =   29
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Middle :"
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
         Left            =   5160
         TabIndex        =   28
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Last Name :"
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
         Left            =   9360
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail:"
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
         Left            =   1080
         TabIndex        =   26
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "City :"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   3720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "State/Province :"
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
         Left            =   120
         TabIndex        =   24
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Home No. :"
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
         Left            =   7920
         TabIndex        =   23
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "DOB(dd/mm/yyyy):"
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
         Index           =   3
         Left            =   6960
         TabIndex        =   22
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Occupation:"
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
         Left            =   600
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
      End
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
      Left            =   5280
      TabIndex        =   7
      Top             =   6840
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
      Left            =   11400
      TabIndex        =   6
      Top             =   6840
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
      Left            =   9360
      TabIndex        =   5
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "Delete"
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
      TabIndex        =   4
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add"
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
      Left            =   3240
      TabIndex        =   1
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "MEMBERSHIP PROCESSING"
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
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "membership_processing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsdata As ADODB.Recordset
Private Sub CmdAdd_Click()
framememinfo.Enabled = True
TxtMbId.SetFocus
CmdAdd.Visible = False
CmdEdit.Enabled = False
CmdDel.Enabled = False
CmdExit.Enabled = False
CmdExit.Visible = False
CmdCancel.Visible = True
End Sub

Private Sub CmdCancel_Click()
Unload Me
membership_processing.Show
End Sub

Private Sub CmdDel_Click()
Dim memid As Integer
Dim del As Boolean
del = False
On Error GoTo l1
memid = InputBox("Enter Member ID To Delete.")
Call connection.connection

Set rsdata = New ADODB.Recordset
rsdata.Open "select *from membershipprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
If memid = rsdata(0) Then
TxtMbId.Text = rsdata(0)
TxtFname.Text = rsdata(1)
TxtMname.Text = rsdata(2)
TxtLname.Text = rsdata(3)
TxtEmail.Text = rsdata(4)
TxtOccu.Text = rsdata(5)
If rsdata(6).Value = True Then

OptMale.Value = True
Else

OptFemale.Value = True
End If

TxtStreet.Text = rsdata(7)
TxtCity.Text = rsdata(8)
TxtState.Text = rsdata(9)
TxtCell.Text = rsdata(10)
TxtHome.Text = rsdata(11)
DTPDOB.Value = rsdata(12)
conf = MsgBox("Are you sure you want to delete this member's record?", vbYesNo)
    If conf = vbYes Then
    rsdata.Delete
    rsdata.Update
    rsdata.UpdateBatch
    rsdata.Close
    
    del = True
    Exit Do
    Else
    del = False
    
    Exit Do
    End If
Else
rsdata.MoveNext
End If

Loop

 If del = True Then
MsgBox "Member Deleted Successfully."
Unload Me
membership_processing.Show
End If
l1:
End Sub

Private Sub CmdEdit_Click()
Dim memid As Integer
On Error GoTo l2
memid = InputBox("Enter Member ID To Edit.")
Call connection.connection

Set rsdata = New ADODB.Recordset
rsdata.Open "select *from MembershipProcessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF

If memid = rsdata(0) Then
TxtMbId.Text = rsdata(0)
TxtFname.Text = rsdata(1)
TxtMname.Text = rsdata(2)
TxtLname.Text = rsdata(3)
TxtEmail.Text = rsdata(4)
TxtOccu.Text = rsdata(5)
If rsdata(6).Value = True Then

OptMale.Value = True
Else

OptFemale.Value = True
End If

TxtStreet.Text = rsdata(7)
TxtCity.Text = rsdata(8)
TxtState.Text = rsdata(9)
TxtCell.Text = rsdata(10)
TxtHome.Text = rsdata(11)
DTPDOB.Value = rsdata(12)
MsgBox "Now Edit Records And Update."
framememinfo.Enabled = True
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

If TxtMbId.Text = "" Then
MsgBox "No Member Found. Member ID Wrong."
End If
l2:
End Sub

Private Sub Cmdexit_Click()
Unload Me
End Sub

Private Sub CmdPrint_Click()
Dim mid As Integer
Dim found As Boolean
found = False
On Error GoTo l3
mid = InputBox("Enter Member ID To View Or Print.")
Call connection.connection
Set rsdata = New ADODB.Recordset
rsdata.Open "select *from membershipprocessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
If mid = rsdata(0) Then
member_info.Sections("details").Controls("member_ID").Caption = rsdata(0)
member_info.Sections("details").Controls("first_name").Caption = rsdata(1)
member_info.Sections("details").Controls("middle_name").Caption = rsdata(2)
member_info.Sections("details").Controls("last_name").Caption = rsdata(3)
member_info.Sections("details").Controls("email").Caption = rsdata(4)
member_info.Sections("details").Controls("occu").Caption = rsdata(5)
If rsdata(6).Value = True Then
member_info.Sections("details").Controls("gender").Caption = "Male"
Else
member_info.Sections("details").Controls("gender").Caption = "Female"

End If

member_info.Sections("details").Controls("street_name").Caption = rsdata(7)
member_info.Sections("details").Controls("city_name").Caption = rsdata(8)
member_info.Sections("details").Controls("state_name").Caption = rsdata(9)
member_info.Sections("details").Controls("cell_phone").Caption = rsdata(10)
member_info.Sections("details").Controls("home_number").Caption = rsdata(11)
member_info.Sections("details").Controls("dob").Caption = rsdata(12)

member_info.Show
   found = True
    Exit Do
       
Else
rsdata.MoveNext
End If
Loop
If found = False Then MsgBox "Member Id Doesn't Exist."
l3:
End Sub


Private Sub CmdReset_Click()
TxtMbId.Text = ""
TxtFname.Text = ""
TxtMname.Text = ""
TxtLname.Text = ""
TxtEmail.Text = ""
TxtOccu.Text = ""
OptMale.Value = False
OptFemale.Value = False
TxtStreet.Text = ""
TxtCity.Text = ""
TxtState.Text = ""
TxtCell.Text = ""
TxtHome.Text = ""
DTPDOB.Value = 1 / 1 / 2000


End Sub

Private Sub CmdSave_Click()
Call connection.connection

Set rsdata = New ADODB.Recordset
rsdata.Open "select *from MembershipProcessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
rsdata.MoveNext
Loop

rsdata.AddNew
rsdata(0) = TxtMbId.Text
rsdata(1) = TxtFname.Text
rsdata(2) = TxtMname.Text
rsdata(3) = TxtLname.Text
rsdata(4) = TxtEmail.Text
rsdata(5) = TxtOccu.Text
If OptMale.Value = True Then
rsdata(6).Value = True
ElseIf OptFemale.Value = True Then
rsdata(6).Value = True
End If
rsdata(7) = TxtStreet.Text
rsdata(8) = TxtCity.Text
rsdata(9) = TxtState.Text
rsdata(10) = TxtCell.Text
rsdata(11) = TxtHome.Text
rsdata(12) = DTPDOB.Value
rsdata.Update
rsdata.UpdateBatch
rsdata.Close
conn.Close
MsgBox "Data Saved Successfully."
Unload Me
membership_processing.Show
End Sub

Private Sub CmdUpdate_Click()
rsdata(0) = TxtMbId.Text
rsdata(1) = TxtFname.Text
rsdata(2) = TxtMname.Text
rsdata(3) = TxtLname.Text
rsdata(4) = TxtEmail.Text
rsdata(5) = TxtOccu.Text
If OptMale.Value = True Then
rsdata(6).Value = True
ElseIf OptFemale.Value = True Then
rsdata(6).Value = True
End If
rsdata(7) = TxtStreet.Text
rsdata(8) = TxtCity.Text
rsdata(9) = TxtState.Text
rsdata(10) = TxtCell.Text
rsdata(11) = TxtHome.Text
rsdata(12) = DTPDOB.Value
rsdata.Update
rsdata.UpdateBatch

rsdata.Close
MsgBox "Record Updated Successfully."
Unload Me
membership_processing.Show


End Sub

