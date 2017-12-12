VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form book_return 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Book Return"
   ClientHeight    =   9390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12450
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   12450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   2880
      TabIndex        =   5
      Top             =   840
      Width           =   3015
   End
   Begin VB.TextBox TxtAccss 
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
      Left            =   2880
      TabIndex        =   4
      Top             =   1560
      Width           =   3015
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
      Left            =   4680
      TabIndex        =   3
      Top             =   3600
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
      Left            =   2760
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "Return"
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
      Left            =   840
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form4.frx":0000
      Height          =   5175
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   23
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DataMember      =   "Command1"
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "MemberIDNo"
         Caption         =   "MemberIDNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "FirstName"
         Caption         =   "FirstName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Middle"
         Caption         =   "Middle"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "LastName"
         Caption         =   "LastName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "AccessionNo"
         Caption         =   "AccessionNo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "TitleName"
         Caption         =   "TitleName"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "IssueDate"
         Caption         =   "IssueDate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Returneddate"
         Caption         =   "Returneddate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1679.811
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Member ID No. :"
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
      Left            =   480
      TabIndex        =   11
      Top             =   840
      Width           =   2055
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
      Index           =   1
      Left            =   480
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Member's Name :"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label LblMbname 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   8
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Name :"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblbookname 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   6
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "BOOK RETURN"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "book_return"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdReset_Click()
    TxtMbId.SetFocus
    TxtMbId = ""
    TxtAccss = ""
    LblMbname.Caption = ""
    lblbookname.Caption = ""
   
End Sub

Private Sub CmdReturn_Click()
    Dim find As Boolean
        find = False
    If TxtMbId.Text = "" And Not TxtAccss.Text = "" Then
        MsgBox "Member ID Field Is Empty."
        TxtMbId.SetFocus
    ElseIf TxtAccss = "" And Not TxtMbId.Text = "" Then
        MsgBox "Accession No. Field Is Empty."
        TxtAccss.SetFocus
    ElseIf TxtMbId.Text = "" And TxtAccss.Text = "" Then
        MsgBox "Both Fields Are Empty."
        TxtMbId.SetFocus
    ElseIf lblbookname.Caption = "" And Not LblMbname.Caption = "" Then
        MsgBox "Book Accession No. Is Invalid."
    ElseIf LblMbname.Caption = "" And Not lblbookname.Caption = "" Then
        MsgBox "Member ID Is Invalid."
    ElseIf LblMbname.Caption = "" And lblbookname.Caption = "" Then
        MsgBox "Member ID And Accession No. Are Invalids."
    Else

        Call connection.connection
    
        Set rsdata = New ADODB.Recordset
        rsdata.Open "select *from booktransaction", conn, adOpenDynamic, adLockBatchOptimistic
    Do Until rsdata.EOF
        If TxtMbId.Text = rsdata(0) And TxtAccss.Text = rsdata(1) Then
        
        rsdata(3) = Date
        
        rsdata.Update
        rsdata.UpdateBatch
        rsdata.Close
        MsgBox "Book Returned Success."
        TxtMbId.SetFocus
        find = True
        Exit Do
        
       
        Else
        
        rsdata.MoveNext
        End If
    
    Loop
    If find = True Then
    
     Unload Me
     book_return.Show
     
    Else
        MsgBox "Book Unissued Or Returned Already."
     End If
     
    End If
    
End Sub



Private Sub TxtAccss_Change()
    lblbookname.Caption = ""

Call connection.connection

    Set rsdata = New ADODB.Recordset
    rsdata.Open "select *from bookProcessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
    accss = Val(TxtAccss.Text)
    If accss = rsdata(2) Then
    lblbookname.Caption = rsdata(3)

Exit Do

Else
    rsdata.MoveNext
End If

Loop
End Sub

Private Sub TxtMbId_Change()
    LblMbname.Caption = ""

Call connection.connection

Set rsdata = New ADODB.Recordset
    rsdata.Open "select *from MembershipProcessing", conn, adOpenDynamic, adLockBatchOptimistic
Do Until rsdata.EOF
    mem = Val(TxtMbId.Text)
    If mem = rsdata(0) Then
    LblMbname.Caption = rsdata(1) + " " + rsdata(2) + " " + rsdata(3)

Exit Do

Else
    rsdata.MoveNext
End If

Loop
End Sub
