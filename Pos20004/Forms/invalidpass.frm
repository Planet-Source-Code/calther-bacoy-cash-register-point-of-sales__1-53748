VERSION 5.00
Begin VB.Form invalidpass 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   780
      TabIndex        =   0
      Top             =   780
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800000&
      BorderWidth     =   4
      Height          =   1185
      Left            =   45
      Top             =   45
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   180
      Picture         =   "invalidpass.frx":0000
      Stretch         =   -1  'True
      Top             =   270
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Invalid Password "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   780
      TabIndex        =   1
      Top             =   330
      Width           =   1605
   End
End
Attribute VB_Name = "invalidpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SendKeys "{Home}+{End}"
Password.txtpassword.SetFocus
Unload Me
End Sub
