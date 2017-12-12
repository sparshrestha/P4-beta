VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13950
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   2
      ItemData        =   "Form2.frx":0000
      Left            =   12240
      List            =   "Form2.frx":0028
      TabIndex        =   34
      Top             =   5040
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   1
      ItemData        =   "Form2.frx":0068
      Left            =   10920
      List            =   "Form2.frx":0090
      TabIndex        =   33
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      ItemData        =   "Form2.frx":00D0
      Left            =   9960
      List            =   "Form2.frx":0131
      TabIndex        =   32
      Top             =   5040
      Width           =   855
   End
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
      Index           =   5
      Left            =   9480
      TabIndex        =   30
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Index           =   4
      Left            =   11520
      TabIndex        =   29
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Index           =   3
      Left            =   7440
      TabIndex        =   28
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Index           =   2
      Left            =   5400
      TabIndex        =   27
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
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
      Index           =   1
      Left            =   3360
      TabIndex        =   26
      Top             =   6240
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
      Index           =   0
      Left            =   1200
      TabIndex        =   25
      Top             =   6240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   9
      Left            =   9960
      TabIndex        =   24
      Top             =   4320
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   8
      Left            =   9960
      TabIndex        =   23
      Top             =   3600
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   7
      Left            =   2760
      TabIndex        =   21
      Top             =   5040
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   6
      Left            =   2760
      TabIndex        =   19
      Top             =   4320
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   5
      Left            =   2760
      TabIndex        =   17
      Top             =   3600
      Width           =   3975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Third"
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
      Left            =   12600
      TabIndex        =   16
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Female"
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
      Left            =   11160
      TabIndex        =   15
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Male"
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
      Left            =   9960
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   4
      Left            =   2760
      TabIndex        =   13
      Top             =   2880
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   2
      Left            =   11400
      TabIndex        =   12
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   11
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   10
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   9
      Top             =   1440
      Width           =   3375
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
      Left            =   7080
      TabIndex        =   31
      Top             =   5040
      Width           =   2415
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
      Left            =   8040
      TabIndex        =   22
      Top             =   4320
      Width           =   1335
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
      Left            =   240
      TabIndex        =   20
      Top             =   5040
      Width           =   1935
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
      Left            =   1560
      TabIndex        =   18
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail :"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   2880
      Width           =   975
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
      Left            =   9600
      TabIndex        =   7
      Top             =   2160
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
      Left            =   5400
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Cellphone :"
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
      Left            =   8040
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Street :"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3600
      Width           =   975
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
      Left            =   8280
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "First Name :"
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
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
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
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
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
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
