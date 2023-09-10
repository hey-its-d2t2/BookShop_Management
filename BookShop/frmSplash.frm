VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSplash 
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   4995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   240
      Top             =   240
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin MSComctlLib.ProgressBar ProgressBar1 
         DragMode        =   1  'Automatic
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   3840
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1920
         TabIndex        =   4
         Top             =   4080
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   600
         TabIndex        =   3
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "Kora Kagaz"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   705
         Left            =   2400
         TabIndex        =   2
         Top             =   3000
         Width           =   2760
      End
      Begin VB.Image Image1 
         Height          =   2520
         Left            =   2520
         Picture         =   "frmSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2520
      End
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
Timer1.Enabled = False
Unload Me
LoginForm.Show
End If
End Sub
