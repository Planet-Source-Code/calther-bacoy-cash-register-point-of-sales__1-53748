VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Products 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Product Entry"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13785
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4395
   ScaleWidth      =   13785
   Begin VB.TextBox txtProduct 
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
      Index           =   0
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   23
      Top             =   1020
      Width           =   1200
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   1
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   1740
      Width           =   2400
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   2
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   21
      Top             =   1380
      Width           =   2400
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   3
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   20
      Top             =   2100
      Width           =   960
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   4
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   15
      TabIndex        =   19
      Top             =   2460
      Width           =   1890
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   5
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2820
      Width           =   1890
   End
   Begin VB.TextBox txtProduct 
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
      Index           =   6
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   17
      Top             =   3210
      Width           =   1020
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
      Left            =   5160
      TabIndex        =   15
      Top             =   3960
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of All Products"
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
      Height          =   3045
      Left            =   4410
      TabIndex        =   12
      Top             =   810
      Width           =   9315
      Begin MSComctlLib.ListView lv1 
         Height          =   2715
         Left            =   90
         TabIndex        =   13
         ToolTipText     =   "doouble click to edit or delete records"
         Top             =   210
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4789
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
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
         MouseIcon       =   "Products.frx":0000
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Category"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Quantity"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit Price"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Sell Price"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Reorder Pt."
            Object.Width           =   1852
         EndProperty
      End
   End
   Begin LVbuttons.LaVolpeButton btnSave 
      Height          =   375
      Left            =   3960
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
      MICON           =   "Products.frx":031A
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
      Left            =   5160
      TabIndex        =   10
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
      MICON           =   "Products.frx":0336
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
      Left            =   2760
      TabIndex        =   0
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
      MICON           =   "Products.frx":0352
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
      Left            =   6360
      TabIndex        =   11
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
      MICON           =   "Products.frx":036E
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
      Left            =   7560
      TabIndex        =   1
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
      MICON           =   "Products.frx":038A
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
   Begin LVbuttons.LaVolpeButton btnsearch 
      Height          =   315
      Left            =   3840
      TabIndex        =   24
      Top             =   1380
      Width           =   345
      _ExtentX        =   609
      _ExtentY        =   556
      BTYPE           =   6
      TX              =   "..."
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
      FCOL            =   0
      FCOLO           =   0
      EMBOSSM         =   12632256
      EMBOSSS         =   16777215
      MPTR            =   0
      MICON           =   "Products.frx":03A6
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
      Left            =   4470
      TabIndex        =   16
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label lblcount 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   11610
      TabIndex        =   14
      Top             =   3990
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000009&
      FillColor       =   &H00FFFFFF&
      Height          =   15
      Index           =   0
      Left            =   2760
      Top             =   540
      Width           =   6030
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   15
      Index           =   0
      Left            =   2760
      Top             =   540
      Width           =   6030
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   9
      Top             =   2100
      Width           =   780
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   2460
      Width           =   825
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selling Price  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   2820
      Width           =   990
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reorder Point  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product ID  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description  :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1740
      Width           =   945
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category Type :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1380
      Width           =   1185
   End
End
Attribute VB_Name = "Products"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim autoProd As New ADODB.Recordset
Dim Prod As New ADODB.Recordset
Dim ls As ListItem
Private Sub btnAdd_Click()
Call autoProdId
Call unlocking
Call clearing
btnSave.Enabled = True
btnCancel.Enabled = True
btnAdd.Enabled = False
End Sub
Private Sub btnCancel_Click()
btnAdd.Enabled = True
btnSave.Enabled = False
btnDelete.Enabled = False
btnEdit.Enabled = False
btnEdit.Caption = "&Edit"
Call Cleartext(txtProduct, 7)
End Sub
Private Sub btnDelete_Click()
Dim rsDelete As New ADODB.Recordset
If rsDelete.State = 1 Then Set rsDelete = Nothing

rsDelete.Open "SELECT * from [Product] where [Product Id]='" & txtProduct(0).Text & "'", con, adOpenDynamic, adLockPessimistic

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
                    Call clearing
                    
                End If
        
        Else
            MsgBox "This Record is already Deleted!", vbExclamation + vbOKOnly
                    Call clearing
        End If
    
    End With

End Sub
Private Sub btnEdit_Click()
If btnEdit.Caption = "&Edit" Then
   btnEdit.Caption = "&Update"
   btnCancel.Enabled = True
   Call unlocking
   btnAdd.Enabled = False
   btnSave.Enabled = False
   btnDelete.Enabled = False
Else: btnEdit.Caption = "&Update"
btnEdit.Caption = "&Edit"

Dim rsProd As New ADODB.Recordset

If rsProd.State = 1 Then Set rsProd = Nothing

rsProd.Open "SELECT * from [Product] where [Product ID] ='" & txtProduct(0).Text & "'", con, adOpenDynamic, adLockPessimistic
   
   With rsProd
       
     If Not TxtBoxIsEmpty(txtProduct, 7) Then
        
        con.BeginTrans
     
        For i = 0 To 6
            
            .Fields(i) = txtProduct(i).Text
            
        Next i
        
        .Update
        .Requery
        
        con.CommitTrans
        
        .Close
        Call txtSearch_Change
        btnEdit.Enabled = False
        btnAdd.Enabled = True
      Else
            MsgBox "Please Complete Data!", vbExclamation + vbOKCancel
      
      End If
     
   End With
   
Set rsProd = Nothing

End If
End Sub

Private Sub btnSave_Click()
Dim rsProd As New ADODB.Recordset

If rsProd.State = 1 Then Set rsProd = Nothing

rsProd.Open "SELECT * from [Product] where [Product ID] ='" & txtProduct(0).Text & "'", con, adOpenDynamic, adLockPessimistic
   
   With rsProd
       
     If .EOF Then
   
        con.BeginTrans
        .AddNew
        For i = 0 To 6
            
            .Fields(i) = txtProduct(i).Text
            
        Next i
        .Update
        .Requery
        con.CommitTrans
        .Close
        Call txtSearch_Change
        Call clearing
        Call locking
        btnAdd.Enabled = True
        btnSave.Enabled = False
      Else
      
            MsgBox "Duplicate Found Can't Save", vbInformation + vbOKOnly
        
      End If
   
   End With
   
Set rsProd = Nothing
End Sub
Private Sub Form_Load()

If Prod.State = 1 Then Set Prod = Nothing

    Prod.Open "SELECT * from Product", con, adOpenDynamic, adLockPessimistic
                     dpProd
                            
Me.Width = 13905
Me.Height = 4800
Me.Top = 0
End Sub
Function autoProdId()

Randomize
txtProduct(0).Text = Round(Rnd() * 999999) & txtProduct(0).Text + Chr(Round(Rnd() * 25) + 65)

End Function
Function locking()
For i = 0 To 6
    txtProduct(i).locked = True
Next i
End Function
Function unlocking()
For i = 0 To 6
    txtProduct(i).locked = False
Next i
End Function
Function clearing()
For i = 1 To 6
    txtProduct(i).Text = ""
    Next i
End Function
Private Sub btnsearch_Click()
CategoryList.Show
End Sub

Private Sub lv1_DblClick()
btnEdit.Enabled = True
btnDelete.Enabled = True
btnCancel.Enabled = True
btnSave.Enabled = False
txtProduct(0).Text = lv1.SelectedItem.Text
txtProduct(1).Text = lv1.SelectedItem.SubItems(1)
txtProduct(2).Text = lv1.SelectedItem.SubItems(2)
txtProduct(3).Text = lv1.SelectedItem.SubItems(3)
txtProduct(4).Text = lv1.SelectedItem.SubItems(4)
txtProduct(5).Text = lv1.SelectedItem.SubItems(5)
txtProduct(6).Text = lv1.SelectedItem.SubItems(6)
End Sub
Private Sub txtProduct_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

    If Index = 1 Then
        txtProduct(1).Text = UCase(txtProduct(1).Text)
        txtProduct(3).SetFocus
    End If
    
    If Index = 3 Then
        SendKeys "{Home}+{End}"
        txtProduct(4).SetFocus
    End If
        
    If Index = 4 Then
        txtProduct(4).Text = CStr(Format(txtProduct(4).Text, "####.#0"))
        SendKeys "{Home}+{End}"
        txtProduct(5).SetFocus
    End If
    
    If Index = 5 Then
        txtProduct(5).Text = CStr(Format(txtProduct(5).Text, "####.#0"))
        SendKeys "{Home}+{End}"
        txtProduct(6).SetFocus
    End If
    
    If Index = 6 Then
       If Not btnSave.Enabled = False Then
       End If
    End If


    
End If
End Sub
Private Sub txtProduct_KeyPress(Index As Integer, KeyAscii As Integer)

If Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then

    Select Case KeyAscii
    
        Case Asc(0) To Asc(9)
        Case Str("8")
        
    Case Else
            KeyAscii = 0
    End Select

End If
End Sub
Function dpProd()
Do While Not Prod.EOF

    Set ls = lv1.ListItems.Add(, , Prod.Fields(0))
        
        ls.SubItems(1) = Prod.Fields(1)
        ls.SubItems(2) = Prod.Fields(2)
        ls.SubItems(3) = Prod.Fields(3)
        ls.SubItems(4) = Prod.Fields(4)
        ls.SubItems(5) = Prod.Fields(5)
        ls.SubItems(6) = Prod.Fields(6)
        Prod.MoveNext
    

Loop

    lblcount.Caption = Str(Prod.RecordCount) + " record(s) found"
Set Prod = Nothing
End Function
Private Sub txtSearch_Change()

If Prod.State = 1 Then Set Prod = Nothing

calther = "SELECT * from [Product] where [Description] like '%" & Trim(txtSearch) & "%'"

            Prod.Open calther, con, adOpenKeyset, adLockOptimistic
                
                
                   lv1.ListItems.Clear

                        dpProd
End Sub
