VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form LookUp 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Product List"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows Default
   Begin LVbuttons.LaVolpeButton btnOk 
      Height          =   405
      Left            =   4230
      TabIndex        =   1
      Top             =   2730
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   714
      BTYPE           =   6
      TX              =   "OK"
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
      MICON           =   "LookUp.frx":0000
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
   Begin MSComctlLib.ListView lv1 
      Height          =   2535
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product Id"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Category Name"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "LookUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Prods As New ADODB.Recordset
Dim lst As ListItem
Private Sub btnOk_Click()
Unload Me
End Sub
Private Sub Form_Load()

   Prods.Open "SELECT [Product Id],[Description],[Category Id] FROM Product order by [Product Id]", con, adOpenDynamic, adLockPessimistic
                loadLvw
End Sub
Function loadLvw()

    
                         

    lv1.ListItems.Clear

        With Prods
        
            Do While Not .EOF
            
                Set lst = lv1.ListItems.Add(, , .Fields(0))
                    
                    lst.SubItems(1) = .Fields(1)
                    lst.SubItems(2) = .Fields(2)
                    .MoveNext
            Loop
        
        End With
        
    Set Prods = Nothing

End Function
