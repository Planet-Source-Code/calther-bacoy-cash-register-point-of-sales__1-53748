VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Category 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Category"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   ForeColor       =   &H8000000A&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   6000
   Begin MSComctlLib.ListView lv 
      Height          =   2355
      Left            =   90
      TabIndex        =   11
      Top             =   1530
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   4154
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category ID"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category Name"
         Object.Width           =   7320
      EndProperty
   End
   Begin VB.TextBox txtSearch 
      DataField       =   "User Name"
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
      Left            =   840
      TabIndex        =   9
      Top             =   3960
      Width           =   1890
   End
   Begin VB.TextBox txtCatname 
      DataField       =   "User Name"
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
      Left            =   1530
      TabIndex        =   1
      Top             =   1050
      Width           =   3090
   End
   Begin VB.TextBox txtCatId 
      DataField       =   "User Name"
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
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   1890
   End
   Begin LVbuttons.LaVolpeButton btnSave 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Save"
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
      MICON           =   "Category.frx":0000
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
   Begin LVbuttons.LaVolpeButton btnEdit 
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   0
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
      MICON           =   "Category.frx":001C
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
      Left            =   0
      TabIndex        =   4
      Top             =   0
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
      MICON           =   "Category.frx":0038
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
      Left            =   3600
      TabIndex        =   5
      Top             =   0
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
      MICON           =   "Category.frx":0054
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
   Begin LVbuttons.LaVolpeButton btnCancel 
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      BTYPE           =   6
      TX              =   "&Cancel"
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
      MICON           =   "Category.frx":0070
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   10
      Top             =   3960
      Width           =   615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   0
      Top             =   480
      Width           =   6030
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   15
      Left            =   0
      Top             =   480
      Width           =   6030
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category Name  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   8
      Top             =   1020
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category ID  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   7
      Top             =   660
      Width           =   1095
   End
End
Attribute VB_Name = "Category"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCategory As New ADODB.Recordset
Dim autoID As New ADODB.Recordset
Dim ls As ListItem
Function CategoryID()
Randomize
txtCatId.Text = "CAT" & Round(Rnd() * 999999) & txtCatId.Text + Chr(Round(Rnd() * 25) + 65)

End Function
Private Sub btnAdd_Click()
Call CategoryID
btnSave.Enabled = True
btnCancel.Enabled = True
btnAdd.Enabled = False
txtCatname.SetFocus
End Sub
Private Sub btnCancel_Click()
On Error Resume Next
adoCategory.CancelUpdate
Call Cancel
End Sub
Private Sub btnDelete_Click()
Dim rsDelete As New ADODB.Recordset
If rsDelete.State = 1 Then Set rsDelete = Nothing

rsDelete.Open "SELECT * from [Category] where [Category Id]='" & txtCatId.Text & "'", con, adOpenDynamic, adLockPessimistic

    With rsDelete
    
        If Not .EOF Then
        
            ans = MsgBox("Are you sure do you want delete this one record?", vbCritical + vbYesNo, "Delete?")
            
                If ans = vbYes Then
                
                    .Delete
                    .Requery
                    .Close
                    Call txtSearch_Change
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
If btnEdit.Caption = "&Edit" Then
   btnEdit.Caption = "&Update"
   btnCancel.Enabled = True
   btnAdd.Enabled = False
   btnSave.Enabled = False
   btnDelete.Enabled = False
Else: btnEdit.Caption = "&Update"
btnEdit.Caption = "&Edit"

Dim rsProd As New ADODB.Recordset

If rsProd.State = 1 Then Set rsProd = Nothing

rsProd.Open "SELECT * from [Category] where [Category ID] ='" & txtCatId.Text & "'", con, adOpenDynamic, adLockPessimistic
   
   With rsProd
       
        
        con.BeginTrans
     
        .Fields(0) = txtCatId.Text
        .Fields(1) = UCase(txtCatname.Text)
        .Update
        .Requery
        
        con.CommitTrans
        
        .Close
        Call txtSearch_Change
        btnEdit.Enabled = False
        btnAdd.Enabled = True
      
     
   End With
   
Set rsProd = Nothing

End If
End Sub

Private Sub btnSave_Click()



If txtCatname.Text = "" Then

    MsgBox "Please fill up Category Name!", vbInformation + vbOKOnly

Else

    
    With adoCategory
    
     
        con.BeginTrans
        .AddNew
        .Fields(0) = txtCatId.Text
        .Fields(1) = UCase(txtCatname.Text)
        .Update
        .Requery
        con.CommitTrans
        Call Cancel
        Call txtSearch_Change
        
           
          
      
      
    End With
End If



End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub Form_Load()


 If adoCategory.State = 1 Then Set adoCategory = Nothing
        
        adoCategory.CursorLocation = adUseClient
        adoCategory.Open "SELECT * FROM [Category] ORDER BY [Category Id]", con, adOpenDynamic, adLockOptimistic
        adoCategory.Requery
        
               dview
        

Call colums
End Sub
Function Cancel()
txtCatId.Text = ""
txtCatname.Text = ""
btnAdd.Enabled = True
btnSave.Enabled = False
btnDelete.Enabled = False
End Function
Function colums()
    Me.Width = 6120
    Me.Height = 4935
End Function
Private Sub lv_Click()
btnEdit.Enabled = True
btnDelete.Enabled = True
txtCatId.Text = lv.SelectedItem.Text
txtCatname.Text = lv.SelectedItem.SubItems(1)
End Sub
Private Sub txtSearch_Change()


If adoCategory.State = 1 Then Set adoCategory = Nothing


calther = "SELECT * from [Category] where [Category Name] like '" & Trim(txtSearch) & "%'"

            adoCategory.Open calther, con, adOpenKeyset, adLockOptimistic
                
                
                   lv.ListItems.Clear
                   
                        dview
              
End Sub
Private Sub dview()

Do While Not adoCategory.EOF

    Set ls = lv.ListItems.Add(, , adoCategory.Fields(0))
        ls.SubItems(1) = adoCategory.Fields(1)
        adoCategory.MoveNext
Loop

End Sub
