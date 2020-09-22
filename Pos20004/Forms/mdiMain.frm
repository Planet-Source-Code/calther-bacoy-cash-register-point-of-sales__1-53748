VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "ABAN's Point of Sales"
   ClientHeight    =   1830
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6315
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   810
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1284
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":17FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":234C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3028
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":347C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":3D58
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":4634
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":5288
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":55A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":6568
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   1244
      ButtonWidth     =   1349
      ButtonHeight    =   1191
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Sales"
            Object.ToolTipText     =   "Sales Entry"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Product"
            Object.ToolTipText     =   "Product Entry"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keycat"
                  Text            =   "&Category Entry"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyprod"
                  Text            =   "&Product Database"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Users"
            Object.ToolTipText     =   "Users Setup"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keyprods"
                  Text            =   "Prod&uct Report"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "keysales"
                  Text            =   "Sales Rep&ort"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Log Off"
            Object.ToolTipText     =   "Log Off"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Exit"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&About"
            Object.ToolTipText     =   "About the Author"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   1455
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:44 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "5/13/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
            Picture         =   "mdiMain.frx":69BA
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   13229
            MinWidth        =   13229
            Text            =   "Copyright(c) Bacoy Software Company. Developed by CALTHER L. BACOY"
            TextSave        =   "Copyright(c) Bacoy Software Company. Developed by CALTHER L. BACOY"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Press F9 to Log Off"
            TextSave        =   "Press F9 to Log Off"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnucat 
         Caption         =   "&Category"
      End
      Begin VB.Menu mnuprod 
         Caption         =   "&Product Database"
      End
      Begin VB.Menu mnusales 
         Caption         =   "&Sales Entry"
      End
      Begin VB.Menu mnuser 
         Caption         =   "Add &Users"
      End
      Begin VB.Menu c 
         Caption         =   "&Log Off"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnurep 
      Caption         =   "&Report"
      Begin VB.Menu mnurepPR 
         Caption         =   "Prod&uct Report"
      End
      Begin VB.Menu mnurepSr 
         Caption         =   "Sales Rep&ort"
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub c_Click()
ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
If ans = vbYes Then
    Password.Show
    Me.Hide
Else
    Exit Sub
End If
End Sub
Private Sub MDIForm_Load()
mnufile.Visible = False
mnurep.Visible = False
End Sub
Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mnufile
End If
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
ans = MsgBox("Are you sure do you want to exit?", vbCritical + vbYesNo, "Exit?")

    If ans = vbYes Then
            
            End
    Else
            Cancel = Not Me
        
    End If
    
End Sub

Private Sub mnucat_Click()
Category.Show
End Sub

Private Sub mnuexit_Click()
Unload Me
End Sub

Private Sub mnuprod_Click()
Products.Show
End Sub

Private Sub mnurepPR_Click()
Dim rs As New ADODB.Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Product", con, adOpenDynamic, adLockPessimistic

    Set drpt1.DataSource = rs
        Me.Toolbar1.Visible = False
        drpt1.Title = "As of " & Format(Date, "mmmm dd, yyyy")
        drpt1.Show
End Sub

Private Sub mnurepSr_Click()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset

With rs1
    .Open "Select * from SalesRep", con, adOpenDynamic, adLockPessimistic
        
        Do While Not .EOF
            With rs2
                .Open "Select * from [Temp] WHERE [Product Id] = '" & rs1.Fields("Product Id") & "'", con, adOpenDynamic, adLockPessimistic
                    If .EOF = True Then
                        con.BeginTrans
                        .AddNew
                        .Fields(0) = rs1.Fields("Product Id")
                        .Fields(1) = rs1.Fields!QtySold
                        .Fields(2) = rs1.Fields!TAmount
                        .Update
                        con.CommitTrans
                    Else
                        con.BeginTrans
                        .Fields!Totalamount = .Fields!Totalamount + rs1.Fields!TAmount
                        .Fields!QtySold = .Fields!QtySold + rs1.Fields!QtySold
                        .Update
                        con.CommitTrans
                    End If
                .Close
                Set rs2 = Nothing
            End With
            .MoveNext
        Loop
    .Close
    Set rs1 = Nothing
End With


    Dim rvd As New ADODB.Recordset
        
        rvd.CursorLocation = adUseClient
        rvd.Open "SELECT * FROM Sales_Report", con, adOpenDynamic, adLockPessimistic
    
        Set drptDailySales.DataSource = rvd
            drptDailySales.Title = "As of " & Format(Date, "mmmm dd, yyyy")
            drptDailySales.Show
            
                mdiMain.Toolbar1.Visible = False
                  
        If rs3.State = 1 Then Set rs3 = Nothing
        
            rs3.Open "Delete * from Temp", con, adOpenDynamic, adLockPessimistic
        
        
End Sub
Private Sub mnusales_Click()
Sales.Show
sb1.Visible = False
End Sub
Private Sub mnuser_Click()
User.Show
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rsSecurity As New ADODB.Recordset
Dim level As String

    If rsSecurity.State = 1 Then Set rsSecurity = Nothing
    
        rsSecurity.Open "SELECT * from [User] where [User Name] ='" & sb1.Panels(5).Text & "'", con, adOpenDynamic, adLockPessimistic
    
        If Not rsSecurity.EOF Then
            level = rsSecurity.Fields("User Level")
        Else
            Exit Sub
        End If
Select Case Button.Index

    Case 1:
            Toolbar1.Visible = False
            Me.Hide
            Sales.Show
    Case 4:
        If level = "Cashier" Then
            MsgBox "You are restricted in this module!", vbCritical, "Restricted"
                    Exit Sub
                    
        Else
            User.Show
        End If
            
    Case 6:
            Call c_Click
    Case 7:
            Unload Me
    Case 8:
    
            AboutMe.Show

End Select
End Sub
Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim rsSecurity As New ADODB.Recordset
Dim level As String

    If rsSecurity.State = 1 Then Set rsSecurity = Nothing
    
        rsSecurity.Open "SELECT * from [User] where [User Name] ='" & sb1.Panels(5).Text & "'", con, adOpenDynamic, adLockPessimistic
    
            level = rsSecurity.Fields("User Level")

Select Case ButtonMenu.Key
    
    Case "keyprod"
        
        If level = "Cashier" Then
            MsgBox "You are restricted in this module!", vbCritical, "Restricted"
                    Exit Sub
                    
        Else
            Products.Show
        End If
    Case "keycat"
        If level = "Cashier" Then
            MsgBox "You are restricted in this module!", vbCritical, "Restricted"
                    Exit Sub
                    
        Else
              Category.Show
        End If
      
    Case "keysales"
    
            Call mnurepSr_Click
       
    Case "keyprods"
            Call mnurepPR_Click
End Select
End Sub
Function prodReport()
Dim rs As New ADODB.Recordset
Dim TQty As String


 
If rs.State = 1 Then Set rs = Nothing

    rs.Open "select * from Product", _
    con, adOpenKeyset, adLockPessimistic
    
    TQty = rs.Fields!Quantity
        
        With drptProduct
            Set .DataSource = Nothing
                .DataMember = ""
            Set .DataSource = rs.DataSource
            
            With .Sections("Section1").Controls
                For i = 1 To .Count
                    If TypeOf .Item(i) Is RptTextBox Then
                        .Item(i).DataMember = ""
                        .Item(i).DataField = rs.Fields(i - 1).Name
                    End If
                Next i
               
                
            End With
             
              .Show
             
                mdiMain.Toolbar1.Visible = False
                
        End With
       

End Function
 
