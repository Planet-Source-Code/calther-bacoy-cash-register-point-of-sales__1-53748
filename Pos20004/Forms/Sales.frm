VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{899348F9-A53A-4D9E-9438-F97F0E81E2DB}#1.0#0"; "LVbuttons.ocx"
Begin VB.Form Sales 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Cash Register"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Height          =   3825
      Left            =   7140
      TabIndex        =   34
      Top             =   5160
      Width           =   7155
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   2700
         TabIndex        =   43
         Top             =   1020
         Width           =   4095
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   300
         Width           =   4080
      End
      Begin VB.TextBox txtPayments 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   585
         Left            =   2700
         MaxLength       =   13
         TabIndex        =   1
         Top             =   1680
         Width           =   4110
      End
      Begin VB.TextBox txtChange 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404080&
         Height          =   525
         Left            =   2700
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "0.00"
         Top             =   2400
         Width           =   4110
      End
      Begin LVbuttons.LaVolpeButton btnCompute 
         Height          =   495
         Left            =   5670
         TabIndex        =   35
         Top             =   3090
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&Compute"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":0000
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
      Begin LVbuttons.LaVolpeButton btnReset 
         Height          =   495
         Left            =   4410
         TabIndex        =   42
         Top             =   3090
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&Reset"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":001C
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
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   270
         Width           =   2235
      End
      Begin VB.Label lblFieldLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   225
         TabIndex        =   40
         Top             =   1005
         Width           =   1545
      End
      Begin VB.Label lblFieldLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Pay :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   210
         TabIndex        =   39
         Top             =   1650
         Width           =   1920
      End
      Begin VB.Label lblFieldLabel 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Change  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   2340
         Width           =   1350
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1365
      Left            =   750
      TabIndex        =   28
      Top             =   7620
      Width           =   6345
      Begin LVbuttons.LaVolpeButton btnSave 
         Height          =   825
         Left            =   1350
         TabIndex        =   29
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
         BTYPE           =   6
         TX              =   "&Save Transaction"
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
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":0038
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   1
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnPrint 
         Height          =   825
         Left            =   2550
         TabIndex        =   30
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
         BTYPE           =   6
         TX              =   "&Print"
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
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":0054
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "5"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnCancel 
         Height          =   825
         Left            =   3750
         TabIndex        =   31
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
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
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":0070
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "4"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnStart 
         Height          =   825
         Left            =   150
         TabIndex        =   32
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
         BTYPE           =   6
         TX              =   "&Start Transaction"
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
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":008C
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "(None)"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   1
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
      Begin LVbuttons.LaVolpeButton btnExit 
         Height          =   825
         Left            =   4950
         TabIndex        =   33
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1455
         BTYPE           =   6
         TX              =   "&Exit"
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
         BCOL            =   14737632
         FCOL            =   8388608
         FCOLO           =   0
         EMBOSSM         =   12632256
         EMBOSSS         =   16777215
         MPTR            =   0
         MICON           =   "Sales.frx":00A8
         ALIGN           =   1
         IMGLST          =   "(None)"
         IMGICON         =   "2"
         ICONAlign       =   0
         ORIENT          =   0
         STYLE           =   1
         IconSize        =   4
         SHOWF           =   -1  'True
         BSTYLE          =   0
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Height          =   2475
      Left            =   750
      TabIndex        =   19
      Top             =   5160
      Width           =   6345
      Begin VB.TextBox txtInvoiceNo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   210
         TabIndex        =   25
         Top             =   1770
         Width           =   2520
      End
      Begin VB.TextBox txtInvoiceDate 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   4170
         TabIndex        =   22
         Top             =   1335
         Width           =   1920
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   480
         Left            =   4170
         TabIndex        =   21
         Top             =   1845
         Width           =   1920
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sales No  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   1290
         Width           =   1530
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   3060
         TabIndex        =   24
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Time  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   6
         Left            =   3060
         TabIndex        =   23
         Top             =   1815
         Width           =   990
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "ABAN'S CELL PHONE TRADING CENTRE ---------------------- Cashier -----------------------"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   840
         Left            =   90
         TabIndex        =   20
         Top             =   450
         Width           =   6120
      End
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Left            =   6360
      TabIndex        =   18
      Top             =   10260
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   291
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   450
      Top             =   810
   End
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   135
      Left            =   5190
      TabIndex        =   17
      Top             =   10260
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Max             =   30
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   4035
      Left            =   750
      TabIndex        =   5
      Top             =   1140
      Width           =   4425
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Enter Barcode"
         Top             =   450
         Width           =   3150
      End
      Begin VB.TextBox txtQtyHand 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
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
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   2160
         Width           =   915
      End
      Begin VB.TextBox txtQtyOrder 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         MaxLength       =   3
         TabIndex        =   10
         ToolTipText     =   "input quantity here"
         Top             =   2520
         Width           =   1140
      End
      Begin VB.TextBox txtPrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1830
         Width           =   1440
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1440
         Width           =   2760
      End
      Begin VB.CommandButton cmdOrder 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Order"
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
         Left            =   2670
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2520
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
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
         Left            =   1500
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2550
         TabIndex        =   55
         Top             =   3750
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F4"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   3
         Left            =   2640
         TabIndex        =   54
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   1710
         TabIndex        =   53
         Top             =   3750
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   2
         Left            =   1800
         TabIndex        =   52
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "START"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   51
         Top             =   3750
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   1
         Left            =   120
         TabIndex        =   50
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   0
         Left            =   960
         TabIndex        =   49
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "F9"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   3480
         TabIndex        =   48
         Top             =   3120
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOOK UP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   870
         TabIndex        =   47
         Top             =   3750
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "LOG OFF"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3390
         TabIndex        =   46
         Top             =   3720
         Width           =   1005
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Product Code/Bar Code/SKU"
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
         Index           =   7
         Left            =   105
         TabIndex        =   27
         Top             =   150
         Width           =   2475
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Qty On Hand :"
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
         Left            =   150
         TabIndex        =   16
         Top             =   2160
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity Order  :"
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
         Left            =   150
         TabIndex        =   14
         Top             =   2490
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price  :"
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
         Left            =   150
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name  :"
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
         Left            =   150
         TabIndex        =   12
         Top             =   1410
         Width           =   1155
      End
      Begin VB.Label Label9 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   1020
         Width           =   915
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "List of Orders"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   4065
      Left            =   5250
      TabIndex        =   3
      Top             =   1110
      Width           =   9045
      Begin MSComctlLib.ListView lvprod 
         Height          =   3765
         Left            =   120
         TabIndex        =   4
         ToolTipText     =   "double click to remove  product in the list"
         Top             =   210
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   6641
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Product Id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Item"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descripiton"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Category Name"
            Object.Width           =   4851
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2293
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   30
      Top             =   810
   End
   Begin MSComctlLib.StatusBar sb2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9825
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   27342
            MinWidth        =   27342
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCashier 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   705
      Left            =   840
      TabIndex        =   45
      Top             =   210
      Width           =   7275
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   780
      TabIndex        =   44
      Top             =   9420
      Width           =   6495
   End
End
Attribute VB_Name = "Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Orders As New ADODB.Recordset
Dim lst As ListItem
Private Sub btnCancel_Click()
If lvprod.ListItems.Count = 0 Then
        btnStart.Enabled = True
        btnPrint.Enabled = False
        btnSave.Enabled = False
        btnCancel.Enabled = False
        Call Cleartext
        txtAmount.Text = ""
        txtDiscount.Text = ""
        txtInvoiceNo.Text = ""
            Exit Sub
Else
ans = MsgBox("Are you sure do you wanto cancel?", vbYesNo + vbQuestion, "Cancel?")

If ans = vbYes Then
        btnStart.Enabled = True
        btnPrint.Enabled = False
        btnSave.Enabled = False
        btnCancel.Enabled = False
        Call Cleartext
        txtAmount.Text = ""
        txtDiscount.Text = ""
        txtInvoiceNo.Text = ""
        DoEvents
        

            
            Bar1.Max = lvprod.ListItems.Count * 10
            
            Do While lvprod.ListItems.Count > 0
                
                Bar1.Visible = True
                        
                        lvprod.ListItems.Remove lvprod.SelectedItem.Index
                            
                            lvprod.Refresh
                        
                            Screen.MousePointer = 1
                            
                           Do While Bar1.Value < Bar1.Max
                                
                                Bar1.Value = Bar1.Value + 1
                                    DoEvents
                            Loop
                        DoEvents
            Loop
                     Bar1.Value = 0
                     Screen.MousePointer = 1
                    Bar1.Visible = False
        End If
        DoEvents
End If

End Sub

Private Sub btnCompute_Click()


             a = Val(txtAmount.Text) * Val(txtDiscount.Text)
             b = Val(txtAmount.Text) - a
             txtAmount.Text = b
        
        If txtDiscount.Text = "" Then
            txtDiscount.Text = "0.00"
        End If
        

        If Val(txtPayments.Text) < Val(txtAmount.Text) Then
           MsgBox "Insufficient Amount " & Chr(10) & "Please change the value " & Chr(10) & "that is greater than or equal " & StrConv(txtAmount, vbUpperCase), vbOKOnly + vbInformation, "Insufficient"
                txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))
        Else
            
          If Val(txtDiscount.Text) < Val(txtAmount.Text) Then
            
                bayad = Val(txtPayments.Text)
                    sukli = bayad - Val(txtAmount.Text)
                        txtChange.Text = CStr(Format(sukli, "##,####0.00"))
                            txtAmount.Text = CStr(Format(txtAmount.Text, "##,####0.00"))
                                 txtPayments.Text = CStr(Format(txtPayments.Text, "##,####0.00"))
                            btnSave.Enabled = True
                        txtPayments.locked = True
                        btnCompute.Enabled = False
                    btnSave.SetFocus
                txtDiscount.locked = True
          Else
            MsgBox "Discount should be less than " & vbCrLf & "the total amount", vbExclamation, "Discount"
          End If
        End If
        
     txtPayments.SetFocus
End Sub
Private Sub btnCompute_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub
Private Sub btnExit_Click()
End
End Sub
Private Sub btnPrint_Click()
On Error Resume Next
Dim tammnt As String, cash As String, _
dscount As String, change As String, _
SNo As String

Dim rs As New ADODB.Recordset
 
If rs.State = 1 Then Set rs = Nothing

    rs.Open "select * from [SalesRep] where [Invoice No]='" & _
    Sales.txtInvoiceNo.Text & "'", con, adOpenKeyset, adLockOptimistic
        
        With drptSales
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
    
     
 If rs.State = 1 Then Set rs = Nothing
                
      rs.Open "select * from [Invoice] where [Invoice No]='" & Sales.txtInvoiceNo.Text & "'", con, adOpenKeyset, adLockOptimistic
           
        tamnt = rs.Fields!Amount
        cash = rs.Fields!Payments
        dscount = rs.Fields!Discount
        change = rs.Fields!change
        SNo = rs.Fields("Invoice No")
        
        With drptSales
       
                .Sections("SDetails").Controls("lbltotal").Caption = Format(tamnt, "###,####0.00")
                .Sections("SDetails").Controls("lblcash").Caption = Format(cash, "###,####0.00")
                .Sections("SDetails").Controls("lbldiscount").Caption = dscount
                .Sections("SDetails").Controls("lblchange").Caption = change
                .Sections("header").Controls("invoiceno").Caption = SNo
                
        End With
     
     Set rs = Nothing

mdiMain.Toolbar1.Visible = False
End Sub
Private Sub btnReset_Click()
txtPayments.Text = ""
    txtChange.Text = ""
        txtDiscount.Text = ""
            txtAmount.Text = Format(txtAmount, "########0.00")
                btnCompute.Enabled = True
                    txtPayments.locked = False
                        txtDiscount.locked = False
                            txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))
                                txtPayments.SetFocus
End Sub
Private Sub btnSave_Click()
Dim adoOrders As New ADODB.Recordset
Dim adoInvoice As New ADODB.Recordset

If lvprod.ListItems.Count <= 0 Then
        Exit Sub
Else
                        
    If txtPayments.Text = Empty Then

    MsgBox "Please enter amount!", vbInformation
        txtPayments.SetFocus
            Exit Sub
    Else
                    
                    If adoOrders.State = 1 Then Set adoOrders = Nothing
                    
                      adoOrders.Open "SELECT * from [Invoice] where [Invoice No] = '" & txtInvoiceNo.Text & "'", con, adOpenDynamic, adLockPessimistic
                    
                          With adoOrders
                                          
                              
                              If .EOF Then
                                  Timer2 = True
                                  con.BeginTrans
                                  .AddNew
                                  .Fields(0) = txtInvoiceNo.Text
                                  .Fields(1) = txtAmount.Text
                                  .Fields(2) = txtDiscount.Text
                                  .Fields(3) = txtPayments.Text
                                  .Fields(4) = txtChange.Text
                                  .Fields(5) = Date
                                  .Update
                                  .Requery
                                  con.CommitTrans
                                  .Close
                              End If
                          End With
                    
                    
                     
                    For i = 1 To lvprod.ListItems.Count
                    
                    If adoInvoice.State = 1 Then Set adoInvoice = Nothing
                    
                      adoInvoice.Open "SELECT * from SIDetails", con, adOpenDynamic, adLockPessimistic
                      
                          With adoInvoice
                          
                                  con.BeginTrans
                                  .AddNew
                                  .Fields(0) = txtInvoiceNo.Text
                                  .Fields(1) = lvprod.ListItems(i).Text
                                  .Fields(2) = lvprod.ListItems(i).SubItems(1)
                                  .Fields(3) = lvprod.ListItems(i).SubItems(4)
                                  .Update
                                  .Requery
                                  con.CommitTrans
                                  .Close
                                  
                          End With
                          
                                  
                                  calther = "update Product set Quantity =  Quantity - '" & lvprod.ListItems(i).SubItems(1) & "' where [Product Id] ='" & lvprod.ListItems(i).Text & "'"
                                                     
                                                     con.Execute calther
                    
                                  
                          
                    Next i
                    
                    btnSave.Enabled = False
                    btnPrint.Enabled = True
    End If

End If
End Sub
Private Sub btnStart_Click()
Call autoSalesNo
txtSearch.locked = False
txtPayments.locked = False
btnStart.Enabled = False
btnCancel.Enabled = True
txtSearch.SetFocus
txtAmount.Text = ""
txtDiscount.Text = ""
End Sub
Private Sub cmdOrder_Click()
Dim ilan As Double

    

txtDiscount.Text = "0.00"
If Val(txtQtyOrder.Text) > Val(txtQtyHand.Text) Then
      MsgBox "The stocks on Product Entry " + vbCrLf + " is not enough pls..change" + vbCrLf + " the value of the item ordered", vbOKOnly + vbInformation, "Insufficient"
           SendKeys "{home}" + "{end}"
            Exit Sub
Else
    If txtID.Text = Empty Or txtQtyOrder.Text = Empty Then
    
        Exit Sub
    
    Else
        
      If Val(txtQtyHand.Text) <= 0 Then
      
         MsgBox "There is no stocks on Product Entry " + vbCrLf + " please re-order to continue sales", vbOKOnly + vbInformation, "Empty Stock"
      
      Else
        ans = MsgBox("Sure to Order this?", vbQuestion + vbYesNo, "Order?")
        
        If ans = vbYes Then
            ilan = Val(txtPrice) * Val(txtQtyOrder)
            txtAmount.Text = CStr(Format(ilan, "##,###.#0"))
            Call orderItem
            Call Cleartext
        Else
           Exit Sub
        End If
      End If
     End If
End If
 txtPayments.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub

Private Sub Form_Load()
Me.Width = 9675
Me.Height = 9315
Timer2 = False
sb2.Visible = False
lblCashier.BorderStyle = 0
lblCashier.Caption = "Your Cashier is: " + mdiMain.sb1.Panels(5).Text
mdiMain.sb1.Visible = False
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
If lvprod.ListItems.Count > 0 Then
    MsgBox "Active transaction detected" + vbCrLf + "please cancel or continue transaction" + vbCrLf + "to close this form", vbExclamation, "Active Trans"
           Cancel = Not Me
End If
End Sub
Private Sub lv1_Click()
   txtSearch.Text = lv1.SelectedItem.SubItems(1)
   txtQtyOrder.SetFocus
End Sub
Private Sub lvprod_DblClick()
If lvprod.ListItems.Count = 0 Then
    MsgBox "No items to remove!", vbOKOnly + vbInformation, "Remove"
Else
    If MsgBox("Are you sure you want to remove  " & Chr(10) & Chr(10) & StrConv(lvprod.SelectedItem.SubItems(2), vbUpperCase), vbYesNo + vbQuestion, "Remove Item") = vbYes Then
              
              lvprod.ListItems.Remove lvprod.SelectedItem.Index
                txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))

                            
    Else
            Exit Sub
    End If
End If

End Sub

Private Sub lvprod_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub

Private Sub Timer1_Timer()
txtInvoiceDate.Text = Format(Date, "mm.dd.yyyy")
txtTime.Text = Time

End Sub
Private Sub Timer2_Timer()
On Error Resume Next
Bar.Value = Bar.Value + 1

Screen.MousePointer = vbHourglass

If Bar.Value <= 10 Then
lblStatus.Caption = "Processing...."
ElseIf Bar.Value <= 20 Then
lblStatus.Caption = "Please wait...."
ElseIf Bar.Value <= 30 Then
lblStatus.Caption = "Transaction Complete"

If Bar.Value = 30 Then
If Timer2.Interval >= 1 Then
lblStatus.Caption = ""
btnPrint.SetFocus
Call btnPrint_Click
Timer2 = False
Screen.MousePointer = vbDefault
End If
End If
End If


End Sub
Private Sub txtDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    
    
    If Val(txtDiscount.Text) < Val(txtAmount.Text) Then
        txtDiscount.Text = txtDiscount.Text / 100
           btnCompute.Enabled = True
               a = Val(txtAmount.Text) * Val(txtDiscount.Text)
                    b = Val(txtAmount.Text) - a
                       txtAmount.Text = b
                           txtPayments.SetFocus
    Else
        MsgBox "Discount should be less than " & vbCrLf & "the total amount", vbExclamation, "Discount"
    End If

End If

End Sub
Private Sub txtPayments_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub

Private Sub txtPayments_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    
    Call btnCompute_Click
    
End If

End Sub
Private Sub txtQtyOrder_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub

Private Sub txtQtyOrder_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    
        Case Asc(0) To Asc(9)
        Case Str("8")
        Case Str("13")
Case Else
          KeyAscii = 0
End Select

If KeyAscii = 13 Then
Call cmdOrder_Click
End If
End Sub
Private Sub txtSearch_Change()

If Orders.State = 1 Then Set Orders = Nothing

calther = "SELECT * from [Product] where [Product Id] like '" & Trim(txtSearch) & "%'"

            Orders.Open calther, con, adOpenKeyset, adLockOptimistic
                
               
            
            With Orders
                  
              If Not .EOF Then
              
                txtID.Text = .Fields(0)
                txtName.Text = .Fields(1)
                txtPrice.Text = CStr(Format(.Fields(5), "##,###.#0"))
                txtQtyHand.Text = .Fields(3)
                
              Else
                Call Cleartext
              End If
            End With
        
End Sub
Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    
    Case vbKeyF1
    
        Call btnStart_Click

    Case vbKeyF2
        
        LookUp.Show vbModal
    Case vbKeyF3
        btnCancel_Click
    Case vbKeyF4
        Call btnSave_Click
    Case vbKeyF9
        ans = MsgBox("Are you sure do you want Log Off?", vbQuestion + vbYesNo, "Log Off")
            If ans = vbYes Then
                Unload Me
                Password.Show
                mdiMain.Hide
            Else
                Exit Sub
            End If
End Select
End Sub
Private Sub txtSearch_KeyPress(KeyAscii As Integer)

If txtSearch.Text = Empty Then
    Exit Sub
Else
    If KeyAscii = 13 Then
        
        txtQtyOrder.SetFocus
        txtPrice.Text = Format(txtPrice.Text, "##,###.#0")
    
    
        
        If Orders.EOF Then
    
            MsgBox "Product not found", vbInformation, "Not Found"
                   Call Cleartext
                   SendKeys "{home}+{end}"
                  txtSearch.SetFocus
                    Exit Sub
        Else
            Call txtSearch_Change
                txtQtyOrder.SetFocus
        End If
    End If
End If
End Sub
Private Sub orderItem()
                 
             Qty = Val(txtQtyOrder.Text)
             
                 Set lst = lvprod.FindItem(txtID.Text, , , lvwPartial)
                    If lst Is Nothing Then
                        Set lst = lvprod.ListItems.Add(, , txtID.Text, , 0)
                         With lst
                                .SubItems(1) = getquantity
                                .SubItems(2) = txtName.Text
                                .SubItems(3) = Orders.Fields("category Id")
                                .SubItems(4) = Format(getquantity * txtPrice.Text, "#,###.#0")
                                txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))
                                txtDiscount.Text = "0.00"
                               
                                
                          End With
                    Else
                            
                               With lst
                               
                                .SubItems(1) = getquantity + .SubItems(1)
                                .SubItems(4) = Format(.SubItems(1) * txtPrice.Text, "#,###.#0")
                                txtAmount.Text = CStr(Format(ComputeAmount, "########0.00"))
                                txtDiscount.Text = "0.00"
                               
                               
                               End With
                    End If

End Sub
Function ComputeAmount() As String
    Dim X As Long
    Dim total As Double

    For X = 1 To lvprod.ListItems.Count
    
        total = Val(total) + lvprod.ListItems(X).SubItems(4)
        
    Next X
    ComputeAmount = CStr(total)
End Function
Function Cleartext()
txtID.Text = ""
txtName.Text = ""
txtPrice.Text = ""
txtQtyOrder.Text = ""
txtQtyHand.Text = ""
txtPayments.Text = ""
txtChange.Text = ""
End Function
Function autoSalesNo()
Randomize
txtInvoiceNo.Text = "INV" & Round(Rnd() * 999999) & txtInvoiceNo.Text + Chr(Round(Rnd() * 25) + 65)

End Function
Private Sub SelectmField(ByVal text_box As MaskEdBox)
    text_box.SelStart = 0
    text_box.SelLength = Len(text_box.Text)
End Sub
Function Alapa()
MsgBox "Ala pa po ito kasi ala pang Screen Saver", vbInformation
End Function
