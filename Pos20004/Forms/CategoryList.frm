VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CategoryList 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Category List"
   ClientHeight    =   3045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3045
   ScaleWidth      =   4455
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
      Left            =   750
      TabIndex        =   1
      Top             =   2520
      Width           =   1890
   End
   Begin MSDataGridLib.DataGrid dg1 
      Height          =   2355
      Left            =   30
      TabIndex        =   0
      ToolTipText     =   "double click here to select category"
      Top             =   30
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   4154
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
         Name            =   "Tahoma"
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
      Left            =   90
      TabIndex        =   2
      Top             =   2520
      Width           =   615
   End
End
Attribute VB_Name = "CategoryList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adoCategoryList As New ADODB.Recordset
Private Sub dg1_dblClick()
Products.txtProduct(2).Text = dg1.Columns(1)
adoCategoryList.Close
Products.txtProduct(1).SetFocus
Unload Me
End Sub
Private Sub Form_Load()

If adoCategoryList.State = 1 Then Set adoCategoryList = Nothing
   adoCategoryList.CursorLocation = adUseClient
   adoCategoryList.Open "SELECT * from Category", con, adOpenDynamic, adLockPessimistic
    
   Set dg1.DataSource = adoCategoryList
   
   With dg1
    
     .Columns(0).Width = 1200
     .Columns(1).Width = 2600
     Me.Height = 3450
     Me.Width = 4575
   End With
End Sub
