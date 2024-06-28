VERSION 5.00
Begin VB.Form customersale 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Form2"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   -135
   ClientWidth     =   20370
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Customer sale.frx":0000
   ScaleHeight     =   11055
   ScaleMode       =   0  'User
   ScaleWidth      =   22307.65
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   10920
      TabIndex        =   40
      Text            =   "Text6"
      Top             =   9120
      Width           =   2655
   End
   Begin VB.ListBox List12 
      Height          =   1035
      Left            =   15000
      TabIndex        =   39
      Top             =   7920
      Width           =   2655
   End
   Begin VB.ListBox List11 
      Height          =   1035
      Left            =   15000
      TabIndex        =   38
      Top             =   6720
      Width           =   2535
   End
   Begin VB.ListBox List10 
      Height          =   1035
      Left            =   15000
      TabIndex        =   37
      Top             =   5400
      Width           =   2535
   End
   Begin VB.ListBox List9 
      Height          =   1035
      Left            =   15000
      TabIndex        =   36
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ListBox List8 
      Height          =   1035
      Left            =   11880
      TabIndex        =   35
      Top             =   7920
      Width           =   2535
   End
   Begin VB.ListBox List7 
      Height          =   1035
      Left            =   11880
      TabIndex        =   34
      Top             =   6720
      Width           =   2535
   End
   Begin VB.ListBox List6 
      Height          =   1035
      Left            =   11880
      TabIndex        =   32
      Top             =   5400
      Width           =   2535
   End
   Begin VB.ListBox List5 
      Height          =   1035
      Left            =   11880
      TabIndex        =   31
      Top             =   4080
      Width           =   2535
   End
   Begin VB.ListBox List4 
      Height          =   1035
      Left            =   8880
      TabIndex        =   30
      Top             =   7920
      Width           =   2415
   End
   Begin VB.ListBox List3 
      Height          =   1035
      Left            =   8880
      TabIndex        =   29
      Top             =   6720
      Width           =   2415
   End
   Begin VB.ListBox List2 
      Height          =   1035
      Left            =   8880
      TabIndex        =   28
      Top             =   5400
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   8880
      TabIndex        =   27
      Top             =   4080
      Width           =   2415
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H000080FF&
      Caption         =   "Proceed For Invoice"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14640
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9360
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   3015
      Left            =   6720
      TabIndex        =   5
      Top             =   120
      Width           =   12015
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   5400
         TabIndex        =   19
         Text            =   "Text5"
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   5400
         TabIndex        =   18
         Text            =   "Text4"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   5400
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   840
         Width           =   2415
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2040
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1440
         Width           =   1335
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
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Text            =   "-------------Select-------------"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   2400
         TabIndex        =   10
         Text            =   "-------------Select-------------"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2400
         TabIndex        =   9
         Text            =   "-------------Select-------------"
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "quantity"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
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
         Left            =   480
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label9 
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
         Left            =   360
         TabIndex        =   7
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label8 
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
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Aadhar no. :"
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
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Total  Amount"
      Height          =   615
      Left            =   7800
      TabIndex        =   41
      Top             =   9120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "Amount"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   33
      Top             =   8160
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Seed"
      Height          =   375
      Index           =   5
      Left            =   15600
      TabIndex        =   26
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "Pesticides"
      Height          =   375
      Index           =   3
      Left            =   12600
      TabIndex        =   25
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Product"
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   24
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Quantity"
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   23
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Rate"
      Height          =   375
      Index           =   0
      Left            =   7080
      TabIndex        =   22
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Fertilizer"
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   21
      Top             =   3360
      Width           =   615
   End
End
Attribute VB_Name = "customersale"
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
Text1.Text = Clear
Text2.Text = Clear
Text3.Text = Clear
Text4.Text = Clear
Text5.Text = Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
sql = "select fertnm from fertdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select pestnm from pestdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo3.AddItem r.Fields(0)
r.MoveNext
Loop
sql = "select seednm from seeddetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo4.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
X
sql = "select nm from custdetail where adno = " + Text2.Text + ""

 Set r = c.Execute(sql)

   If r.EOF = True Then
    s = MsgBox("record not found are you want to update record", vbExclamation + vbYesNo)
         If s = vbYes Then
         Unload Me
         Load customerdetail
          customerdetail.Show
          customerdetail.Height = 11520
          customerdetail.Width = 20490
        customerdetail.ScaleHeight = 11055
        customerdetail.ScaleWidth = 20370
         Else
         Text2.SetFocus
          End If
 Else
 Text3.Text = r.Fields(0)

 End If
 End If
End Sub

