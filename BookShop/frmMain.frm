VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FAFFFF&
   Caption         =   "kora Kagaz"
   ClientHeight    =   10875
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   17910
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":6988A
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00FCFFFF&
      ForeColor       =   &H80000008&
      Height          =   10875
      Left            =   0
      ScaleHeight     =   10845
      ScaleWidth      =   2775
      TabIndex        =   0
      Top             =   0
      Width           =   2805
      Begin VB.CommandButton cmdLogout 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "Logout"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   9360
         Width           =   2415
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   10200
         Width           =   2415
      End
      Begin VB.CommandButton cmdBooksInStock 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "Books In Stock"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5400
         Width           =   2415
      End
      Begin VB.CommandButton cmdNewBook 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "New Book"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   2415
      End
      Begin VB.CommandButton cmdNewOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "New Order"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3480
         Width           =   2415
      End
      Begin VB.CommandButton cmdSaleBook 
         Appearance      =   0  'Flat
         BackColor       =   &H00AA9D23&
         Caption         =   "Sale Book"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   6360
         Width           =   2415
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Kora Kagaz"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   420
         Left            =   600
         TabIndex        =   1
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   2040
         Left            =   120
         Picture         =   "frmMain.frx":FF86C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   2520
      End
   End
   Begin VB.Menu mnuBook 
      Caption         =   "Book"
      Begin VB.Menu mnuNewBook 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuNewOrder 
         Caption         =   "New Order"
      End
      Begin VB.Menu mnuUpdateBook 
         Caption         =   "Update"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuViewBooks 
         Caption         =   "View Book's"
      End
   End
   Begin VB.Menu mnuSale 
      Caption         =   "Sale"
      Begin VB.Menu mnuSaleBook 
         Caption         =   "Sale Book"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "Report"
      Begin VB.Menu mnuCustomerDetails 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu mnuOrderDetails 
         Caption         =   "Order Details"
      End
      Begin VB.Menu mnuBookInStock 
         Caption         =   "Book In Stock"
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub mnuEditBook_Click()

End Sub

Private Sub cmdBooksInStock_Click()
    frmBookStock.Show
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLogout_Click()
    Unload Me
    LoginForm.Show
End Sub

Private Sub cmdNewBook_Click()
    AddNewBook.Show
End Sub

Private Sub cmdNewOrder_Click()
    frmNewOrder.Show
End Sub

Private Sub cmdSaleBook_Click()
    frmSale.Show
End Sub

Private Sub mnuAdmin_Click()
    frmAdminSetting.Show
End Sub

Private Sub mnuBookInStock_Click()
frmBookStock.Show
End Sub

Private Sub mnuCustomerDetails_Click()
frmCustomerDetails.Show
End Sub

Private Sub mnuNewBook_Click()
    AddNewBook.Show
End Sub

Private Sub mnuNewOrder_Click()
    frmNewOrder.Show
End Sub

Private Sub mnuOrderDetails_Click()
    frmOrderDetails.Show
End Sub

Private Sub mnuSaleBook_Click()
    frmSale.Show
End Sub

Private Sub mnuUpdateBook_Click()
    frmEditBookDetails.Show
End Sub

Private Sub mnuViewBooks_Click()
    frmBookStock.Show
End Sub
