VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Splash 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5655
   ControlBox      =   0   'False
   FillColor       =   &H8000000A&
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1890
      Top             =   750
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4170
      Top             =   840
   End
   Begin MSComctlLib.ProgressBar Bar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   2715
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label lblbar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Loading . . ."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   2280
      Width           =   2925
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Calther Lao Bacoy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Index           =   0
      Left            =   660
      TabIndex        =   3
      Top             =   1770
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Author  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   1230
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Point of Sales"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   5385
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
'On Error Resume Next
Static a As Integer
Static b As Integer
Static c As Integer
Static d As Integer
Static e As Integer
Static f As Integer

a = a + 1
b = b + 1
c = c + 1
d = d + 1
e = e + 1
f = f + 1
Label1.Caption = Mid("ABAN's Point of Sales", 1, e)
If e = 22 Then e = 22
Label3(0).Caption = Mid("Calther Lao Bacoy ", 1, f)

If f = 18 Then f = 18


Bar.Value = Bar.Value + 2

Screen.MousePointer = vbHourglass
If Bar.Value = 20 Then
lblbar.Caption = "Loading . . ."
ElseIf Bar.Value = 60 Then
lblbar.Caption = "Please wait . . ."
ElseIf Bar.Value = 100 Then

If Bar.Value = 100 Then

If Timer1.Interval >= 1 Then
Unload Splash
Load Password
Password.Show
Screen.MousePointer = vbDefault
End If
End If
End If
End Sub
