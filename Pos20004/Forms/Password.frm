VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Password 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Log In"
   ClientHeight    =   3690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3360
   ControlBox      =   0   'False
   FillColor       =   &H80000004&
   ForeColor       =   &H80000004&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   3360
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox txtUserName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   150
      TabIndex        =   0
      Top             =   600
      Width           =   3075
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2790
      Top             =   3720
   End
   Begin VB.TextBox txtDay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2070
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   150
      Width           =   1140
   End
   Begin VB.TextBox txtTime 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   150
      Width           =   1140
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   150
      PasswordChar    =   "Ï"
      TabIndex        =   1
      Top             =   2550
      Width           =   3090
   End
   Begin LVbuttons.LaVolpeButton cmdOk 
      Height          =   375
      Left            =   450
      TabIndex        =   4
      Top             =   3150
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "Log &In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Password.frx":0000
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin LVbuttons.LaVolpeButton Command2 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3150
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      BCOL            =   12632256
      FCOL            =   8388608
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Password.frx":001C
      ALIGN           =   1
      IMGLST          =   "(None)"
      IMGICON         =   "(None)"
      ICONAlign       =   0
      ORIENT          =   0
      STYLE           =   1
      IconSize        =   2
      SHOWF           =   -1  'True
      BSTYLE          =   0
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   1440
      Picture         =   "Password.frx":0038
      Stretch         =   -1  'True
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ctr As Integer
Dim xText
Dim adoPrimaryRS As New ADODB.Recordset
Private Sub cmdOk_Click()
With adoPrimaryRS
        .Requery
        .Find "[User Name] like '" & txtUserName.Text & "'"
        
                 
    If Not .EOF Then

            
        If UCase(txtPassword.Text) = UCase(.Fields("Password")) Then
             
                
            
                If .Fields("User Level") = "Cashier" Then
                        
                    
                    adoPrimaryRS.Close
                    mdiMain.sb1.Panels(5).Text = Password.txtUserName.Text
                    Unload Me
                    Sales.Show
                    mdiMain.Toolbar1.Visible = False
                Else
                    adoPrimaryRS.Close
                    mdiMain.sb1.Panels(5).Text = Password.txtUserName.Text
                    Unload Me
                    mdiMain.Toolbar1.Visible = True
                    mdiMain.Show
                End If
         Else
                ctr = ctr + 1
                If ctr = 4 Then
                invalidpass.Show
                   End
                Else
                    If ctr = 1 Then
                        invalidpass.Show
                        invalidpass.Label1.Top = 160
                        invalidpass.Label1.Caption = "You have 3 tries only" + vbCrLf + "  Invalid Password"
                    Else
                        If ctr = 2 Then
                        invalidpass.Show
                        invalidpass.Label1.Top = 160
                        invalidpass.Label1.Caption = "This is your Second (2) Attempt" + vbCrLf + "  Invalid Password"
                        ElseIf ctr = 3 Then
                        invalidpass.Show
                        invalidpass.Label1.Top = 160
                        invalidpass.Label1.Caption = "This your last Attempt" + vbCrLf + "  Invalid Password"
                        End If
                    End If
                    SendKeys "{Home}+{End}"
                End If
         End If
    
    Else
    
     ans = MsgBox("Please select User", vbOKCancel + vbCritical, "User")
    
       If ans = vbCancel Then
          End
       Else
        txtUserName.ListIndex = 0
       End If
       
    End If
    
End With

End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()

If adoPrimaryRS.State = 1 Then Set adoPrimaryRS = Nothing

       adoPrimaryRS.CursorLocation = adUseClient
       
        adoPrimaryRS.Open "SELECT * FROM [User] ORDER BY [User Name]", con, adOpenDynamic, adLockOptimistic
        
               If adoPrimaryRS.RecordCount = 0 Then
                    Exit Sub
                Else
                    adoPrimaryRS.MoveFirst
                        Do While Not adoPrimaryRS.EOF
                            txtUserName.AddItem IIf(IsNull(adoPrimaryRS("User Name")), "", adoPrimaryRS("User Name"))
                               adoPrimaryRS.MoveNext
                        Loop
                End If
       
End Sub
Private Sub Timer1_Timer()
txtTime.Text = Format(Time, "hh:mm:ss")
txtDay.Text = Format(Date, "mm.dd.yyyy")
End Sub
Private Sub txtpassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdOk_Click
End Sub
Private Sub txtUserName_Click()
txtPassword.SetFocus
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtPassword.SetFocus
End Sub
