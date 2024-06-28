VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Fertilizer 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   11085
   ClientLeft      =   480
   ClientTop       =   540
   ClientWidth     =   20400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "Fertilizer.frx":0000
   ScaleHeight     =   11055
   ScaleMode       =   0  'User
   ScaleWidth      =   20320
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "X"
      Height          =   195
      Left            =   18960
      TabIndex        =   25
      Top             =   4920
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      Height          =   2415
      Left            =   16080
      TabIndex        =   18
      Top             =   4920
      Width           =   3135
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   480
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fertilizer "
      BeginProperty Font 
         Name            =   "Rockwell Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   5280
      TabIndex        =   0
      Top             =   3120
      Width           =   9735
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   4080
         TabIndex        =   28
         Top             =   3600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Format          =   120455169
         CurrentDate     =   44445
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   4080
         TabIndex        =   27
         Top             =   3000
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   873
         _Version        =   393216
         Format          =   120455169
         CurrentDate     =   44445
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   4080
         TabIndex        =   24
         Top             =   600
         Width           =   3135
      End
      Begin VB.TextBox Text9 
         Height          =   525
         Left            =   4080
         TabIndex        =   23
         Top             =   5400
         Width           =   3135
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   4080
         TabIndex        =   22
         Top             =   4800
         Width           =   3135
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   4080
         TabIndex        =   21
         Top             =   4200
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5640
         TabIndex        =   17
         Top             =   6480
         Width           =   2415
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Search"
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
         Left            =   2040
         TabIndex        =   16
         Top             =   6360
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Delete"
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
         Left            =   7560
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   15
         Top             =   4200
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Update"
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
         Left            =   7680
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   14
         Top             =   3120
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add"
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
         Left            =   7680
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   13
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "New"
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
         Left            =   7680
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   4080
         TabIndex        =   3
         Top             =   2400
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   4080
         TabIndex        =   2
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   4080
         TabIndex        =   1
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Price/kg :"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   20
         Top             =   5520
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fertilizer Code:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fertilizer name:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   " Company Name:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Batch no.:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Mfg Date:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   7
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Expiry Date"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Quantity/Bag:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   5
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Price/Bag:"
         BeginProperty Font 
            Name            =   "Sitka Small"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   4800
         Width           =   1575
      End
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2895
      Left            =   240
      Picture         =   "Fertilizer.frx":65F1F
      Stretch         =   -1  'True
      Top             =   120
      Width           =   20010
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   7860
      Left            =   240
      Picture         =   "Fertilizer.frx":93F3F
      Stretch         =   -1  'True
      Top             =   3240
      Width           =   4755
   End
End
Attribute VB_Name = "Fertilizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim c As ADODB.Connection
Dim r As ADODB.Recordset
Dim sql As String
Public Function x()
Set c = New ADODB.Connection
c.Open "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=True"
Set r = New ADODB.Recordset

End Function

Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""

Text7.Text = ""
Text8.Text = ""
Text9.Text = ""


Text1.SetFocus
End Sub

Private Sub Command2_Click()
x
sql = " insert into fertdetail values ('" + Text1.Text + "','" + Text2.Text + "','" + Text3.Text + "','" + Text4.Text + "','" + Format(DTPicker1.Value, "dd-mmm-yyyy") + "','" + Format(DTPicker2.Value, "dd-mmm-yyyy") + "'," + Text7.Text + "," + Text8.Text + "," + Text9.Text + ")"

On Error GoTo handleError
Set r = c.Execute(sql)
Combo1.Clear

sql = "select fertcod from fertdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
MsgBox "record saved ", vbInformation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text1.SetFocus
showfertilizer.Visible = False
showfertilizer.Adodc1.Refresh
Exit Sub
handleError:
MsgBox " fertilizer code allready inserted", vbExclamation

Exit Sub


End Sub

Private Sub Command3_Click()
sql = " update  fertdetail set fertnm= '" + Text1.Text + "',fertcod = '" + Text2.Text + "', compnynm = '" + Text3.Text + "',bchno = '" + Text4.Text + "',mfgdt = '" + Format(DTPicker1.Value, "dd-mmm-yyyy") + "',expdt = '" + Format(DTPicker2.Value, "dd-mmm-yyyy") + "',qtybag = " + Text7.Text + ",pcebag = " + Text8.Text + ",pcekg = " + Text9.Text + " where fertcod = '" + Text2.Text + "'"
Set r = c.Execute(sql)
showfertilizer.Adodc1.Refresh
MsgBox "record updated", vbInformation

End Sub

Private Sub Command4_Click()
s = MsgBox(" are you want to delete this record", vbExclamation + vbYesNo)
 If s = vbYes Then
 x
sql = "delete from fertdetail where fertcod = '" + Text2.Text + "' "
c.Execute (sql)
showfertilizer.Adodc1.Refresh
MsgBox "record deleted ", vbExclamation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""

End If
End Sub

Private Sub Command5_Click()
Frame2.Visible = True
Command8.Visible = True

End Sub

Private Sub Command6_Click()
 Unload Fertilizer
 Load showfertilizer
 showfertilizer.Top = 90
 showfertilizer.Left = 10
showfertilizer.Show
showfertilizer.Top = 100
showfertilizer.Height = 11520
showfertilizer.Width = 20490
showfertilizer.ScaleHeight = 11055
showfertilizer.ScaleWidth = 20370
End Sub

Private Sub Command7_Click()
x
sql = "select * from fertdetail where fertcod = '" + Combo1.Text + "'"

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
Text4.Text = r.Fields(3)
End If
If (Not IsNull(r.Fields(4))) Then
DTPicker1 = r.Fields(4)
End If
If (Not IsNull(r.Fields(5))) Then
DTPicker2 = r.Fields(5)
End If
If (Not IsNull(r.Fields(6))) Then
Text7.Text = r.Fields(6)
End If
If (Not IsNull(r.Fields(7))) Then
Text8.Text = r.Fields(7)
End If
If (Not IsNull(r.Fields(8))) Then
Text9.Text = r.Fields(8)
End If

MsgBox "record found", vbExclamation
Command2.Enabled = True
Command3.Enabled = True
Command8.Visible = False
Frame2.Visible = False
End If
End Sub

Private Sub Command8_Click()
Frame2.Visible = False
Command8.Visible = False
End Sub

Private Sub Form_Load()
Text9.Locked = True
Combo1.Clear

Frame2.Visible = False
x

Command8.Visible = False
sql = "select fertcod from fertdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo1.AddItem r.Fields(0)
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
If KeyAscii = 13 Then Text4.SetFocus
End Sub
Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DTPicker1.SetFocus
End Sub

Private Sub Text5_GotFocus()
MonthView1.Visible = True
End Sub



Private Sub Text6_GotFocus()
MonthView2.Visible = True
End Sub



Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
KeyAscii = KeyAscii
 If KeyAscii = 13 Then
 KeyAscii = KeyAscii
 Text8.SetFocus
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
KeyAscii = KeyAscii
 If KeyAscii = 13 Then
 KeyAscii = KeyAscii
 Text9.Text = Text8.Text / Text7.Text
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If

End Sub


Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Text5.Text = MonthView1.Value
MonthView1.Visible = False
Text6.SetFocus

End Sub
Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
Text6.Text = MonthView1.Value
MonthView2.Visible = False
Text7.SetFocus

End Sub

