VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   10500
   ClientLeft      =   -30
   ClientTop       =   765
   ClientWidth     =   19215
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   19155
      TabIndex        =   0
      Top             =   0
      Width           =   19215
      Begin VB.Label Label1 
         Caption         =   "Jai Mahadeo Fertilizer Shop"
         BeginProperty Font 
            Name            =   "Poor Richard"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   6480
         TabIndex        =   1
         Top             =   120
         Width           =   7335
      End
   End
   Begin VB.Menu cust 
      Caption         =   "Customer"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu supp 
      Caption         =   "Supplier"
   End
   Begin VB.Menu product 
      Caption         =   "Products"
      Begin VB.Menu fert 
         Caption         =   "Fertilizer"
      End
      Begin VB.Menu pesti 
         Caption         =   "pesticies"
      End
      Begin VB.Menu seed 
         Caption         =   "seeds"
      End
   End
   Begin VB.Menu sale 
      Caption         =   "Sales"
   End
   Begin VB.Menu purchase 
      Caption         =   "Purchases"
   End
   Begin VB.Menu stock 
      Caption         =   "Stocks"
   End
   Begin VB.Menu report 
      Caption         =   "Reports"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cust_Click()
Unload Me
Load customersale
customersale.Show
End Sub

Private Sub fert_Click()
Unload Me
Load Fertilizer
Fertilizer.Show
End Sub

Private Sub pesti_Click()
Unload Me
Load Pesticides
Pesticides.Show
End Sub

Private Sub purchase_Click()
Unload Me
Load purchases
purchases.Show

End Sub

Private Sub seed_Click()
Unload Me
Load Seeds
Seeds.Show
End Sub

Private Sub supp_Click()
Unload Me
Load supplier
supplier.Show
End Sub
