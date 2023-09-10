VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSale 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "        Sale Book"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   18015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   18015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9240
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DSNBKShopKK"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DSNBKShopKK"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SalesDetails"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FCFFFF&
      Caption         =   "Customer && Book Details"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7575
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   6495
      Begin VB.TextBox txtPay 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   7
         ToolTipText     =   "Enter sale price"
         Top             =   5160
         Width           =   2655
      End
      Begin VB.TextBox txtAuthor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   4
         ToolTipText     =   "Enter author name"
         Top             =   3240
         Width           =   4455
      End
      Begin VB.TextBox txtSalePrice 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   6
         ToolTipText     =   "Enter sale price"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         BackColor       =   &H0091DCF5&
         Caption         =   "&Clear "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Clear Fields "
         Top             =   6480
         Width           =   1815
      End
      Begin VB.TextBox txtBookName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   3
         ToolTipText     =   "Enter book name"
         Top             =   2520
         Width           =   4455
      End
      Begin VB.TextBox txtCustEmail 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   2
         ToolTipText     =   "Enter customer email"
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtCustMob 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   1
         ToolTipText     =   "Enter customer mobile"
         Top             =   1200
         Width           =   4455
      End
      Begin VB.TextBox txtCustName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   0
         ToolTipText     =   "Enter customer name"
         Top             =   600
         Width           =   4455
      End
      Begin VB.CommandButton cmdAdd 
         Appearance      =   0  'Flat
         BackColor       =   &H0091DCF5&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3960
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Add to sale"
         Top             =   6480
         Width           =   1815
      End
      Begin VB.TextBox txtQuantity 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1680
         MaxLength       =   30
         MousePointer    =   3  'I-Beam
         TabIndex        =   5
         ToolTipText     =   "Enter quantity"
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Line Line12 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   4320
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Pay  :  "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   23
         Top             =   5280
         Width           =   705
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   4320
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Price : "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   18
         Top             =   4680
         Width           =   1170
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   4320
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   6120
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   6120
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   6120
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   6120
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity : "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   4000
         Width           =   1110
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Author :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   3320
         Width           =   840
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Book :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   15
         Top             =   2640
         Width           =   645
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Email : "
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   825
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000D&
         X1              =   1680
         X2              =   6120
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblUID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0091DCF5&
         BackStyle       =   0  'Transparent
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   750
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00C0C000&
         BorderStyle     =   3  'Dot
         X1              =   120
         X2              =   6240
         Y1              =   2280
         Y2              =   2280
      End
   End
   Begin VB.CommandButton cmdPrint 
      Appearance      =   0  'Flat
      BackColor       =   &H0091DCF5&
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15240
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Print Bill"
      Top             =   8880
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   7455
      Left            =   7200
      TabIndex        =   22
      ToolTipText     =   "Customer purchase details"
      Top             =   840
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   13150
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Book Details"
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   "Book Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   "Author"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   ""
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   ""
         Caption         =   "Sale Price"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   ""
         Caption         =   "Total Pay"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            WrapText        =   -1  'True
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=DSNBKShopKK"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "DSNBKShopKK"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "CustomerDetails"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sale Book's"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   420
      Left            =   6720
      TabIndex        =   21
      Top             =   0
      Width           =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000D&
      X1              =   6480
      X2              =   8520
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0091DCF5&
      BackStyle       =   0  'Transparent
      Caption         =   "Total :"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   540
      Left            =   9960
      TabIndex        =   20
      Top             =   9000
      Width           =   1110
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H0091DCF5&
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   11280
      TabIndex        =   19
      Top             =   9000
      Width           =   735
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      X1              =   1920
      X2              =   6360
      Y1              =   4320
      Y2              =   4320
   End
End
Attribute VB_Name = "frmSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClearAll()
   txtCustName.Text = ""
    txtCustMob.Text = ""
    txtCustEmail.Text = ""
    txtBookName.Text = ""
    txtAuthor.Text = ""
    txtQuantity.Text = ""
    txtSalePrice.Text = ""
    txtPay.Text = ""
End Sub

Private Sub cmdAdd_Click()
      ' DataGrid2.
   ' DataGrid2 = txtBookName.Text
    'DataGrid2.Columns(1) = txtAuthor.Text
    'DataGrid2.Columns(2) = txtQuantity.Text
    'DataGrid2.Columns(3) = txtSalePrice.Text
    'DataGrid2.Columns(4) = txtPay.Text
    'MSFlexGrid1.AddItem(0) = txtBookName.Text
    'MSFlexGrid1.AddItem(1) = txtAuthor.Text
      'With MSFlexGrid1
     'Row_Lbl.Caption = .Rows
     '.Rows = Val(Row_Lbl.Caption) + 1
     '.TextMatrix(Val(Row_Lbl.Caption) - 1, 0) = Val(Row_Lbl.Caption) - 1
     '.TextMatrix(Val(Row_Lbl.Caption) - 1, 1) = W_Name_Txt.Text
     '.TextMatrix(Val(Row_Lbl.Caption) - 1, 2) = W_Age_Txt.Text
     '.TextMatrix(Val(Row_Lbl.Caption) - 1, 3) = W_Mono_Txt.Text
  'End With
 Clear
End Sub
Private Sub Clear()
    txtAuthor.Text = ""
    txtQuantity.Text = ""
    txtSalePrice.Text = ""
  txtPay.Text = ""
End Sub

Private Sub CustomerDetails()
           Adodc1.Recordset.AddNew
With Adodc1.Recordset
    .Fields(0).Value = txtCustName.Text
    .Fields(1).Value = txtCustMob.Text
    .Fields(2).Value = txtCustEmail.Text
End With
Adodc1.Recordset.Update
Adodc1.Refresh
 ClearAll
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdPrint_Click()
    CustomerDetails
    DataGrid2.ClearFields
End Sub
Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
       If KeyAscii = vbKeyReturn Then
      txtQuantity.SetFocus
    End If
End Sub
Private Sub txtBookName_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
      txtAuthor.SetFocus
    End If
End Sub

Private Sub txtCustEmail_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
      txtBookName.SetFocus
    End If
End Sub

Private Sub txtCustMob_KeyPress(KeyAscii As Integer)
      If KeyAscii = vbKeyReturn Then
      txtCustEmail.SetFocus
    End If
End Sub

Private Sub txtCustName_KeyPress(KeyAscii As Integer)
     If KeyAscii = vbKeyReturn Then
        txtCustMob.SetFocus
    End If
End Sub

Private Sub txtPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdSave.SetFocus
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtSalePrice.SetFocus
    End If
End Sub
Private Sub txtSalePrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPay.Text = Val(txtQuantity.Text() * txtSalePrice.Text())
        txtPay.SetFocus
    End If
End Sub

