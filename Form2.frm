VERSION 5.00
Begin VB.Form freport 
   Caption         =   "Report"
   ClientHeight    =   9975
   ClientLeft      =   720
   ClientTop       =   405
   ClientWidth     =   18945
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   9975
   ScaleWidth      =   18945
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5040
      TabIndex        =   0
      Top             =   3600
      Width           =   8295
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Customer Report"
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
         Left            =   840
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase Report"
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
         Left            =   840
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Sale Report"
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
         Left            =   840
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   2775
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Supplier Report"
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
         Left            =   4680
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Stock Report"
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
         Left            =   4680
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Product Report"
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
         Left            =   4680
         MaskColor       =   &H008080FF&
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3240
         Width           =   2535
      End
   End
   Begin VB.Image Image3 
      Height          =   6555
      Left            =   14040
      Picture         =   "Form2.frx":65F1F
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4035
   End
   Begin VB.Image Image2 
      Height          =   6555
      Left            =   360
      Picture         =   "Form2.frx":76FBA
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   2340
      Left            =   120
      Picture         =   "Form2.frx":811CF
      Stretch         =   -1  'True
      Top             =   240
      Width           =   20400
   End
End
Attribute VB_Name = "freport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load DataReport4

DataReport4.Show
End Sub

Private Sub Command2_Click()
Unload Me
Load DataReport7

DataReport7.Show
End Sub

Private Sub Command3_Click()
Unload Me
Load DataReport6

DataReport6.Show
End Sub

Private Sub Command4_Click()
Unload Me
Load DataReport5

DataReport5.Show
End Sub

Private Sub Command6_Click()
Unload Me
Load freport1
freport1.Show
freport1.Height = 11520
freport1.Width = 20490
freport1.ScaleHeight = 11055
freport1.ScaleWidth = 203701
End Sub
