VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewOrder 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "         New Order"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   17235
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H0091DCF5&
      Caption         =   "&New"
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
      Left            =   360
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "New order "
      Top             =   6600
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   17295
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   10560
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   873
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
         RecordSource    =   "NewOrder"
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmNewOrder.frx":0000
         Height          =   8415
         Left            =   6120
         TabIndex        =   11
         Top             =   960
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   14843
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   22
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
            Name            =   "Palatino Linotype"
            Size            =   9.75
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
               LCID            =   16393
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
               LCID            =   16393
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   8535
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   5655
         Begin VB.TextBox txtAuthor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   450
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   5
            ToolTipText     =   "Enter author name"
            Top             =   3120
            Width           =   3375
         End
         Begin VB.TextBox txtQuantity 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ToolTipText     =   "Enter quantity of books "
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
            Left            =   1800
            MaskColor       =   &H00008000&
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Clear Fields "
            Top             =   7440
            Width           =   1815
         End
         Begin VB.TextBox txtBookName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   450
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   4
            ToolTipText     =   "Enter book name"
            Top             =   2400
            Width           =   3375
         End
         Begin VB.TextBox txtCustEmail 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   450
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   3
            ToolTipText     =   "Enter customer email"
            Top             =   1680
            Width           =   3375
         End
         Begin VB.TextBox txtCustMob 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   450
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   2
            ToolTipText     =   "Enter customer mobile"
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtCustName 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   450
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   1
            ToolTipText     =   "Enter customer name"
            Top             =   480
            Width           =   3375
         End
         Begin VB.CommandButton cmdSave 
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
            Left            =   3240
            MaskColor       =   &H00008000&
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Add to order"
            Top             =   5520
            Width           =   1815
         End
         Begin VB.TextBox txtPublication 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            ToolTipText     =   "Enter publiser name"
            Top             =   3960
            Width           =   3375
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
            Caption         =   "Quantity :"
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
            TabIndex        =   19
            Top             =   4680
            Width           =   1050
         End
         Begin VB.Line Line8 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
            Y1              =   4320
            Y2              =   4320
         End
         Begin VB.Line Line7 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
            Y1              =   3600
            Y2              =   3600
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
            Y1              =   2880
            Y2              =   2880
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0091DCF5&
            BackStyle       =   0  'Transparent
            Caption         =   "Publication :"
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
            Top             =   4005
            Width           =   1275
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   1320
            Width           =   825
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5040
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
            TabIndex        =   13
            Top             =   720
            Width           =   750
         End
         Begin VB.Line Line10 
            BorderColor     =   &H00C0C000&
            BorderStyle     =   3  'Dot
            X1              =   120
            X2              =   5040
            Y1              =   2280
            Y2              =   2280
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "New Order"
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
         Left            =   6240
         TabIndex        =   20
         Top             =   120
         Width           =   1605
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   6000
         X2              =   8040
         Y1              =   600
         Y2              =   600
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      X1              =   1800
      X2              =   6240
      Y1              =   4200
      Y2              =   4200
   End
End
Attribute VB_Name = "frmNewOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub ClearAll()
    txtCustName.Text = ""
    txtCustMob.Text = ""
    txtCustEmail.Text = ""
    txtBookName.Text = ""
    txtAuthor.Text = ""
    txtPublication.Text = ""
    txtQuantity.Text = ""
End Sub

Private Sub cmdClear_Click()
    ClearAll
End Sub

Private Sub cmdNew_Click()
    ClearAll
End Sub

Private Sub cmdSave_Click()
    Adodc1.Recordset.AddNew
With Adodc1.Recordset
    .Fields(0).Value = txtBookName.Text
    .Fields(1).Value = txtAuthor.Text
    .Fields(2).Value = txtPublication.Text
    .Fields(3).Value = txtQuantity.Text
    .Fields(4).Value = txtCustName.Text
    .Fields(5).Value = txtCustMob.Text
    .Fields(6).Value = txtCustEmail.Text
End With
Adodc1.Recordset.Update
Adodc1.Refresh
MsgBox "Record saved.", vbOKOnly + vbInformation, "Information"
Adodc1.Refresh
End Sub



Private Sub txtAuthor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtPublication.SetFocus
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
Private Sub txtPublication_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
      txtQuantity.SetFocus
    End If
End Sub
Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      cmdSave.SetFocus
    End If
End Sub
