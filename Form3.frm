VERSION 5.00
Begin VB.Form freport1 
   Caption         =   "Product Report"
   ClientHeight    =   9510
   ClientLeft      =   -1215
   ClientTop       =   405
   ClientWidth     =   18255
   FillColor       =   &H00C0FFFF&
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   18255
   Begin VB.CommandButton command4 
      BackColor       =   &H008080FF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Pesticides Report"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6120
      Width           =   4575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Seeds Report"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   4575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fertilizer Report"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   6555
      Left            =   600
      Picture         =   "Form3.frx":65F1F
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   4155
   End
   Begin VB.Image Image1 
      Height          =   3420
      Left            =   480
      Picture         =   "Form3.frx":76FBA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   23760
   End
End
Attribute VB_Name = "freport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Load DataReport1

DataReport1.Show
End Sub

Private Sub Command2_Click()
Unload Me
Load DataReport3

DataReport3.Show

End Sub

Private Sub Command3_Click()
Unload Me
Load DataReport2

DataReport2.Show
End Sub

Private Sub Command4_Click()
Unload Me
Load freport
freport.Show
freport.Height = 11520
freport.Width = 20490
freport.ScaleHeight = 11055
freport.ScaleWidth = 20370
End Sub
