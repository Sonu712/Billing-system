VERSION 5.00
Begin VB.Form purchases 
   Caption         =   "Form3"
   ClientHeight    =   8910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16155
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Purchase.frx":0000
   ScaleHeight     =   8910
   ScaleWidth      =   16155
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   255
      Left            =   9480
      TabIndex        =   19
      Top             =   8400
      Width           =   3495
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   14640
      TabIndex        =   18
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Add Supplier"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7680
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Sitka Subheading"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2775
      Left            =   4920
      TabIndex        =   11
      Top             =   4680
      Width           =   3495
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   360
         TabIndex        =   15
         Top             =   1920
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Text            =   "-----------Select--------------"
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "Sitka Subheading"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Supplier Id:"
         BeginProperty Font 
            Name            =   "Sitka Subheading"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Product Order"
      BeginProperty Font 
         Name            =   "Sitka Subheading"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4455
      Left            =   8880
      TabIndex        =   1
      Top             =   3600
      Width           =   5295
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2640
         TabIndex        =   7
         Text            =   "-------------Select-------------"
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2640
         TabIndex        =   6
         Text            =   "-------------Select-------------"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2760
         TabIndex        =   5
         Text            =   "-------------Select-------------"
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Sitka Subheading"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Sitka Subheading"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Sitka Subheading"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fertilizers:-"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   10
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Pesticides:-"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Seeds:-"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   8
         Top             =   3240
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Choose Supplier"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   4920
      TabIndex        =   16
      Top             =   3960
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Jai Mahadeo Fertilizer Shop"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   855
      Left            =   6360
      TabIndex        =   0
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Image Image2 
      Height          =   7005
      Left            =   240
      Picture         =   "Purchase.frx":65F1F
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   120
      Picture         =   "Purchase.frx":70134
      Stretch         =   -1  'True
      Top             =   120
      Width           =   22260
   End
End
Attribute VB_Name = "purchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim c As ADODB.Connection
Dim r As ADODB.Recordset
Dim sql As String
Public Function X()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=True"
Set r = New ADODB.Recordset

End Function
Private Sub Form_Load()
X
sql = "select fertcod from fertdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select pestcod from pestdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select seedcod from seeddetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub
