VERSION 5.00
Begin VB.Form supplier 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   9675
   ClientLeft      =   870
   ClientTop       =   1065
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "supplier.frx":0000
   ScaleHeight     =   9675
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "X"
      Height          =   315
      Left            =   18720
      TabIndex        =   25
      Top             =   4680
      Width           =   375
   End
   Begin VB.Frame Frame2 
      Caption         =   "Search"
      Height          =   3015
      Left            =   16680
      TabIndex        =   20
      Top             =   4680
      Width           =   2535
      Begin VB.CommandButton Command6 
         BackColor       =   &H0000FFFF&
         Caption         =   "Show all"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2280
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0080FFFF&
         Caption         =   "Search"
         Height          =   375
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1560
         Width           =   975
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Text            =   "---Select----"
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Supplier Id:-"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080FFFF&
      Caption         =   "delete"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "save"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Sitka Small"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   14400
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   5400
      TabIndex        =   0
      Top             =   3480
      Width           =   8655
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   3840
         TabIndex        =   13
         Text            =   "--Select--"
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   3840
         TabIndex        =   12
         Top             =   5040
         Width           =   3615
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3840
         TabIndex        =   11
         Top             =   4320
         Width           =   3615
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3840
         TabIndex        =   10
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3840
         TabIndex        =   9
         Top             =   2040
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3840
         TabIndex        =   8
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3840
         TabIndex        =   7
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ifsc Code :"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         TabIndex        =   14
         Top             =   4800
         Width           =   1545
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   8640
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C no. :"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         TabIndex        =   6
         Top             =   4320
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Mob No . :"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "State:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Address:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Id:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   1
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.Image Image2 
      Height          =   6375
      Left            =   360
      Picture         =   "supplier.frx":65F1F
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   3120
      Left            =   120
      Picture         =   "supplier.frx":6CE67
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20010
   End
End
Attribute VB_Name = "supplier"
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
Public Function y()
Set r = New ADODB.Recordset
r.ActiveConnection = c
r.Source = " select * from suppdetail"
r.CursorLocation = adUseClient
r.CursorType = adOpenDynamic
r.LockType = adLockBatchOptimistic
r.Open
Set showsupplier.DataGrid1.DataSource = r
showsupplier.DataGrid1.ReBind
End Function


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
End Sub

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text1.SetFocus
End Sub

Private Sub Command2_Click()
X
sql = " insert into suppdetail values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Combo1.Text + "'," + Text4.Text + "," + Text5.Text + ",'" + Text6.Text + "')"
On Error GoTo handleError
Set r = c.Execute(sql)
MsgBox "record saved ", vbInformation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo2.Clear
sql = "select supp_id from suppdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
y
Exit Sub
handleError:
MsgBox "supplier id allready inserted", vbExclamation

Exit Sub


End Sub
      
Private Sub Command3_Click()
sql = " update  suppdetail set supp_id= '" + Text1.Text + "',supp_nm = '" + Text2.Text + "', supp_add = '" + Text3.Text + "',state ='" + Combo1.Text + "',mob = " + Text4.Text + ",ac_no = " + Text5.Text + ",ifsc = '" + Text6.Text + "' where supp_id = '" + Text1.Text + "'"
Set r = c.Execute(sql)
MsgBox "record updated", vbInformation
y
End Sub

Private Sub Command4_Click()
s = MsgBox(" are you want to delete this record", vbExclamation + vbYesNo)
 If s = vbYes Then
 X
sql = "delete from suppdetail where supp_id = '" + Text1.Text + "' "
c.Execute (sql)
MsgBox "record deleted ", vbExclamation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Clear
y
End If
End Sub

Private Sub Command5_Click()
Frame2.Visible = True
Command8.Visible = True

End Sub

Private Sub Command6_Click()
 Unload supplier
 Load showsupplier
showsupplier.Show

End Sub

Private Sub Command7_Click()
X
sql = "select * from suppdetail where supp_id = '" + Combo2.Text + "'"

 Set r = c.Execute(sql)
 If r.EOF = True Then
 MsgBox "record not found", vbExclamation
 Else
If (Not IsNull(r.Fields(0))) Then
Text1.Text = r.Fields(0)
End If
If (Not IsNull(r.Fields(1))) Then
Text2.Text = r.Fields(1)
End If
If (Not IsNull(r.Fields(2))) Then
Text3.Text = r.Fields(2)
End If
If (Not IsNull(r.Fields(3))) Then
Combo1.Text = r.Fields(3)
End If
If (Not IsNull(r.Fields(4))) Then
Text4.Text = r.Fields(4)
End If
If (Not IsNull(r.Fields(5))) Then
Text5.Text = r.Fields(5)
End If
If (Not IsNull(r.Fields(6))) Then
Text6.Text = r.Fields(6)
End If

End If

MsgBox "record found", vbExclamation

Command8.Visible = False
Frame2.Visible = False
End Sub

Private Sub Command8_Click()
Frame2.Visible = False
Command8.Visible = False
End Sub

Private Sub Form_Load()
Combo1.Clear
Combo2.Clear

Frame2.Visible = False
X

Command8.Visible = False
sql = "select supp_id from suppdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo2.AddItem r.Fields(0)
r.MoveNext
Loop
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Combo1.SetFocus
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
KeyAscii = KeyAscii
 If KeyAscii = 13 Then
 KeyAscii = KeyAscii
 Text5.SetFocus
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
KeyAscii = KeyAscii
 If KeyAscii = 13 Then
 KeyAscii = KeyAscii
Text6.SetFocus
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If

End Sub





