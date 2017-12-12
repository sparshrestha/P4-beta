VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   ClientHeight    =   7470
   ClientLeft      =   5775
   ClientTop       =   1275
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   10365
   Begin VB.CommandButton Command1 
      Caption         =   "RESET"
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
      Index           =   0
      Left            =   7080
      TabIndex        =   20
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   1245
      Left            =   2640
      TabIndex        =   19
      Top             =   4800
      Width           =   6015
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   2640
      List            =   "Form1.frx":0016
      TabIndex        =   17
      Top             =   4200
      Width           =   4935
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   2640
      TabIndex        =   16
      Top             =   3480
      Width           =   4935
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   2640
      TabIndex        =   15
      Top             =   2760
      Width           =   4935
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   2640
      TabIndex        =   14
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   6960
      TabIndex        =   13
      Top             =   1320
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   2640
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command5 
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
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      TabIndex        =   9
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
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
      TabIndex        =   8
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   360
      TabIndex        =   7
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
      TabIndex        =   18
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
      Caption         =   "Author Name/s :"
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
      Caption         =   "Identity No. :"
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
