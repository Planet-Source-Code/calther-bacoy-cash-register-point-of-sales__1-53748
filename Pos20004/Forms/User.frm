VERSION 5.00
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form User 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Add User"
   ClientHeight    =   4515
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   3885
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   2325
      Left            =   60
      TabIndex        =   9
      Top             =   2100
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   4101
      _Version        =   393216
      BackColor       =   14737632
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1140
      Locked          =   -1  'True
      PasswordChar    =   "Ï"
      TabIndex        =   8
      Top             =   600
      Width           =   2010
   End
   Begin VB.ComboBox cbouserlevel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   990
      Width           =   1995
   End
   Begin VB.TextBox txtPassword 
      DataField       =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   255
      Width           =   2010
   End
   Begin LVbuttons.LaVolpeButton btnEdit 
      Height          =   375
      Left            =   1380
      TabIndex        =   5
      Top             =   1620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Edit"
      ENAB            =   0   'False
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
      MICON           =   "User.frx":0000
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
   Begin LVbuttons.LaVolpeButton btnAdd 
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Top             =   1620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Add"
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
      MICON           =   "User.frx":001C
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
   Begin LVbuttons.LaVolpeButton btnDelete 
      Height          =   375
      Left            =   2580
      TabIndex        =   7
      Top             =   1620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Delete"
      ENAB            =   0   'False
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
      MICON           =   "User.frx":0038
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
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Level  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   960
      Width           =   900
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   570
      Width           =   840
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   195
      Width           =   930
   End
End
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoPrimaryRS As New ADODB.Recordset
Private Sub btnAdd_Click()
btnAdd.Enabled = False
btnEdit.Enabled = True
btnEdit.Caption = "&Save"
unlocked
End Sub

Private Sub btnDelete_Click()
Dim rsDelete As New ADODB.Recordset
If rsDelete.State = 1 Then Set rsDelete = Nothing

rsDelete.Open "SELECT * from [User] where [User Name]='" & txtPassword(0).Text & "'", con, adOpenDynamic, adLockPessimistic

    With rsDelete
    
        If Not .EOF Then
        
            ans = MsgBox("Are you sure do you want delete this one record?", vbCritical + vbYesNo, "Delete?")
            
                If ans = vbYes Then
                
                    .Delete
                    .Requery
                    .Close
                    adoPrimaryRS.Requery
                    dg1.Refresh
                    Call dgwidth

                    btnAdd.Enabled = True
                    btnEdit.Enabled = False
                    btnDelete.Enabled = False
                    
                End If
        
        Else
            MsgBox "This Record is already Deleted!", vbExclamation + vbOKOnly
        End If
    
    End With

End Sub
Private Sub btnEdit_Click()
txtPassword(0).locked = True
txtPassword(1).locked = False

If cbouserlevel.Text = "" Or txtPassword(0).Text = "" Or txtPassword(1).Text = "" Then
    
        MsgBox "Please Dont Leave Textbox Empty!", vbInformation, "Aban"
            txtPassword(0).locked = False
            txtPassword(0).SetFocus
            Exit Sub
            
End If
If btnEdit.Caption = "&Edit" Then
    btnEdit.Caption = "&Update"
    btnDelete.Enabled = False
Else:
If btnEdit.Caption = "&Update" Then


 Dim rsPas As New ADODB.Recordset

If rsPas.State = 1 Then Set rsPas = Nothing

rsPas.Open "SELECT * from [User] where [User Name] ='" & txtPassword(0).Text & "'", con, adOpenDynamic, adLockPessimistic
   
   With rsPas
       
     If Not TxtBoxIsEmpty(txtPassword, 2) Then
        
        con.BeginTrans
     
            
        .Fields(0) = txtPassword(0).Text
        .Fields(1) = txtPassword(1).Text
        .Fields(2) = cbouserlevel.Text
        .Update
        .Requery
        adoPrimaryRS.Requery
        dg1.Refresh
        Call dgwidth
        con.CommitTrans
        
        .Close
        btnEdit.Enabled = False
        btnAdd.Enabled = True
       locked
      Else
            MsgBox "Please Complete Data!", vbExclamation + vbOKCancel
      
      End If
     
   End With
   
Set rsProd = Nothing

Else
Dim rsPass As New ADODB.Recordset

If rsPass.State = 1 Then Set rsProd = Nothing

rsPass.Open "SELECT * from [User]", con, adOpenDynamic, adLockPessimistic
   
   With rsPass
       
        con.BeginTrans
        
        .AddNew
        .Fields(0) = txtPassword(0).Text
        .Fields(1) = txtPassword(1).Text
        .Fields(2) = cbouserlevel.Text
        .Update
        .Requery
        adoPrimaryRS.Requery
        dg1.Refresh
        Call dgwidth
        con.CommitTrans
        .Close
        btnAdd.Enabled = True
        btnEdit.Enabled = False
        btnEdit.Caption = "&Edit"
        
       locked
   
   End With
   
Set rsPass = Nothing

End If
End If

End Sub

Private Sub dg1_Click()
txtPassword(0).Text = dg1.Columns(0)
txtPassword(1).Text = dg1.Columns(1)
btnDelete.Enabled = True
btnEdit.Enabled = True
End Sub
Private Sub Form_Load()
 If adoPrimaryRS.State = 1 Then Set adoPrimaryRS = Nothing
        adoPrimaryRS.CursorLocation = adUseClient
        adoPrimaryRS.Open "SELECT * FROM [User] ORDER BY [User Name]", con, adOpenDynamic, adLockOptimistic
        adoPrimaryRS.Requery
        
       
Set dg1.DataSource = adoPrimaryRS

With cbouserlevel
    .AddItem "Admin"
    .AddItem "Cashier"
End With
Me.Width = 4005
Me.Height = 4920

Call dgwidth
End Sub
Function unlocked()
For i = 0 To 1
    txtPassword(i).locked = False
Next
End Function
Function locked()
For i = 0 To 1
    txtPassword(i).locked = True
Next

End Function
Function dgwidth()
With dg1
    .Columns(0).Width = 1200
    .Columns(1).Width = 1000
    .Columns(2).Width = 1200
End With
End Function
