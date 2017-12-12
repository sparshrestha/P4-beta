VERSION 5.00
Begin VB.Form book_processing 
   BackColor       =   &H8000000B&
   Caption         =   "Book Processing"
   ClientHeight    =   7470
   ClientLeft      =   5775
   ClientTop       =   1275
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10365
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
      Left            =   7080
      TabIndex        =   20
      Top             =   6720
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
      Left            =   2040
      TabIndex        =   19
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox TxtBooksyn 
      Height          =   1245
      Left            =   2640
      TabIndex        =   18
      Top             =   4800
      Width           =   6015
   End
   Begin VB.ComboBox CmbClass 
      Height          =   315
      Left            =   2640
      TabIndex        =   16
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox TxtAuthor 
      Height          =   405
      Left            =   2640
      TabIndex        =   15
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox TxtTitle 
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox TxtAcc 
      Height          =   405
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox TxtIsbn 
      Height          =   405
      Left            =   6960
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox TxtBookId 
      Height          =   405
      Left            =   2640
      TabIndex        =   11
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "EXIT"
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
      Left            =   8760
      TabIndex        =   10
      Top             =   6720
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
      Left            =   5400
      TabIndex        =   9
      Top             =   6720
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
      Left            =   2040
      TabIndex        =   7
      Top             =   6720
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
      Left            =   3720
      TabIndex        =   21
      Top             =   6720
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
      Left            =   3720
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
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
      Left            =   0
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
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
      Left            =   3480
      TabIndex        =   6
      Top             =   240
      Width           =   2895
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
      TabIndex        =   5
      Top             =   3480
      Width           =   2055
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
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
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
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
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
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
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
      TabIndex        =   1
      Top             =   2040
      Width           =   1815
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
      TabIndex        =   0
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "book_processing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label7_Click()

End Sub
