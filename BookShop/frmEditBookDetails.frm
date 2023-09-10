VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEditBookDetails 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "         Update Book Details"
   ClientHeight    =   10425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10425
   ScaleWidth      =   17760
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrev 
      Appearance      =   0  'Flat
      BackColor       =   &H0091DCF5&
      Caption         =   "<"
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
      Left            =   600
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Previous record"
      Top             =   8040
      Width           =   1695
   End
   Begin VB.CommandButton cmdNext 
      Appearance      =   0  'Flat
      BackColor       =   &H0091DCF5&
      Caption         =   ">"
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
      Left            =   3600
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Next record"
      Top             =   8040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   17415
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   1800
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
         RecordSource    =   "NewBook"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00FCFFFF&
         Caption         =   "Book Details"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8895
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   5775
         Begin VB.TextBox txtPurchasePrice 
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
            Height          =   480
            Left            =   2040
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   6
            ToolTipText     =   "Enter purchase price"
            Top             =   3480
            Width           =   2655
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
            Height          =   480
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   5
            ToolTipText     =   "Enter quantity"
            Top             =   2880
            Width           =   2655
         End
         Begin VB.CommandButton cmdUpdate 
            Appearance      =   0  'Flat
            BackColor       =   &H0091DCF5&
            Caption         =   "&Update"
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
            Left            =   3360
            MaskColor       =   &H00008000&
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Update book details"
            Top             =   5280
            Width           =   1695
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
            Height          =   480
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   1
            ToolTipText     =   "Enter book name"
            Top             =   480
            Width           =   3495
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
            Height          =   480
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   2
            ToolTipText     =   "Enter book author name"
            Top             =   1080
            Width           =   3495
         End
         Begin VB.TextBox txtPublication 
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
            Height          =   480
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   3
            ToolTipText     =   "Enter book publisher name"
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox txtEdition 
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
            Height          =   480
            Left            =   1680
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   4
            ToolTipText     =   "Enter book edition"
            Top             =   2280
            Width           =   3495
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            BackColor       =   &H0091DCF5&
            Caption         =   "&Delete"
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
            ToolTipText     =   "Delete book"
            Top             =   5280
            Width           =   1695
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
            Height          =   480
            Left            =   2040
            MaxLength       =   30
            MousePointer    =   3  'I-Beam
            TabIndex        =   7
            ToolTipText     =   "Enter selling prioce"
            Top             =   4080
            Width           =   2655
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
            TabIndex        =   20
            Top             =   600
            Width           =   750
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5160
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label2 
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
            TabIndex        =   19
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0091DCF5&
            BackStyle       =   0  'Transparent
            Caption         =   "Publication : "
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
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0091DCF5&
            BackStyle       =   0  'Transparent
            Caption         =   "Edition : "
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
            Top             =   2400
            Width           =   915
         End
         Begin VB.Label Label5 
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
            TabIndex        =   16
            Top             =   3000
            Width           =   1050
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H0091DCF5&
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Price : "
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
            Top             =   3600
            Width           =   1680
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5160
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5160
            Y1              =   2160
            Y2              =   2160
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   5160
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line7 
            BorderColor     =   &H8000000D&
            X1              =   1680
            X2              =   4320
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line8 
            BorderColor     =   &H8000000D&
            X1              =   2040
            X2              =   4680
            Y1              =   3960
            Y2              =   3960
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
            TabIndex        =   14
            Top             =   4320
            Width           =   1170
         End
         Begin VB.Line Line9 
            BorderColor     =   &H8000000D&
            X1              =   2040
            X2              =   4680
            Y1              =   4560
            Y2              =   4560
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmEditBookDetails.frx":0000
         Height          =   8775
         Left            =   6240
         TabIndex        =   12
         ToolTipText     =   "Book's Recoard"
         Top             =   1080
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   15478
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
      Begin VB.Line Line1 
         BorderColor     =   &H8000000D&
         X1              =   6240
         X2              =   9720
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Update Book Detail's"
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
         Left            =   6480
         TabIndex        =   21
         Top             =   240
         Width           =   2985
      End
   End
   Begin VB.Line Line6 
      BorderColor     =   &H8000000D&
      X1              =   2040
      X2              =   6480
      Y1              =   4440
      Y2              =   4440
   End
End
Attribute VB_Name = "frmEditBookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub getData()

    If Adodc1.Recordset.BOF Or Adodc1.Recordset.EOF Then
        If Adodc1.Recordset.BOF Then
               MsgBox "You reached at first data", vbOKOnly + vbInformation, "Information"
        Else
            MsgBox "You reached at last data", vbOKOnly + vbInformation, "Information"
        End If
    Else
    With Adodc1.Recordset
        txtBookName.Text = .Fields(0).Value
        txtAuthor.Text = .Fields(1).Value
        txtPublication.Text = .Fields(2).Value
        txtEdition.Text = .Fields(3).Value
        txtQuantity.Text = .Fields(4).Value
        txtPurchasePrice.Text = .Fields(5).Value
        txtSalePrice.Text = .Fields(6).Value
    End With
    End If
End Sub

Private Sub cmdDelete_Click()
      Dim Reply As Integer
        Reply = MsgBox("Do you want delete.", vbYesNo + vbInformation, "Exit ?")
        If Reply = vbYes Then
            Adodc1.Recordset.Delete
             MsgBox "record deleted.", vbOKOnly + vbInformation, "Information"
        End If
End Sub

Private Sub cmdNext_Click()
    If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.MoveLast
    getData
Else
    Adodc1.Recordset.MoveNext
    getData
End If
End Sub

Private Sub cmdPrev_Click()
If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
        getData
    Else
        Adodc1.Recordset.MovePrevious
         getData
    End If
End Sub

Private Sub cmdUpdate_Click()
    With Adodc1.Recordset
    .Fields(0).Value = txtBookName.Text
    .Fields(1).Value = txtAuthor.Text
    .Fields(2).Value = txtPublication.Text
    .Fields(3).Value = txtEdition.Text
    .Fields(4).Value = txtQuantity.Text
    .Fields(5).Value = txtPurchasePrice.Text
    .Fields(6).Value = txtSalePrice.Text
End With
Adodc1.Recordset.Update
MsgBox "Details Updated.", vbOKOnly + vbInformation, "Information"
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
Private Sub txtEdition_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtQuantity.SetFocus
    End If
End Sub

Private Sub txtPublication_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtEdition.SetFocus
    End If
End Sub
Private Sub txtPurchasePrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      txtSalePrice.SetFocus
    End If
End Sub

Private Sub txtQuantity_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
      txtPurchasePrice.SetFocus
    End If
End Sub
Private Sub txtSalePrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
      cmdUpdate.SetFocus
    End If
End Sub

