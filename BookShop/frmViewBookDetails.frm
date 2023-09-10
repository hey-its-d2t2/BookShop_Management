VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmViewBookDetails 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10380
   ScaleWidth      =   16695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
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
      Left            =   480
      MaskColor       =   &H00008000&
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Delete book entry"
      Top             =   9360
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Caption         =   "Book Details"
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
      Height          =   10095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   16455
      Begin VB.CommandButton Command1 
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
         Left            =   13800
         MaskColor       =   &H00008000&
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print book list"
         Top             =   9240
         Width           =   1815
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   570
         Left            =   15840
         MaskColor       =   &H80000018&
         Picture         =   "frmViewBookDetails.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Close"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   8175
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   15495
         _ExtentX        =   27331
         _ExtentY        =   14420
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   27
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
            Size            =   12
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
   End
End
Attribute VB_Name = "frmViewBookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub
