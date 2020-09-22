VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form AboutMe 
   BackColor       =   &H00E0E0E0&
   Caption         =   "About Me"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7470
   ControlBox      =   0   'False
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   7470
   Begin LVbuttons.LaVolpeButton OK 
      Height          =   375
      Left            =   6060
      TabIndex        =   0
      Top             =   2760
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   14737632
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "AboutMe.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   0
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "63-02-4378119/+639215131799"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4290
      TabIndex        =   6
      Top             =   2010
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   " or contact us  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2010
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "caltherlao@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   1500
      Width           =   1815
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For more Information or Suggestion please email us :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   2790
      TabIndex        =   3
      Top             =   1320
      Width           =   3225
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aban's Cell Phone Tradings Centre Point-of-Sales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   90
      Width           =   4245
      WordWrap        =   -1  'True
   End
   Begin VB.Label MessageList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) 20004 Calther Lao Bacoy -Author"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   2
      Left            =   2760
      TabIndex        =   1
      Top             =   690
      Width           =   4065
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   30
      Picture         =   "AboutMe.frx":001C
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2565
   End
End
Attribute VB_Name = "AboutMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.Width = 7590
   Me.Height = 3660
End Sub
Private Sub OK_Click()
Unload Me
End Sub
