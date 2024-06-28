VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form customerdetail 
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Customer entry.frx":0000
   ScaleHeight     =   11055
   ScaleMode       =   0  'User
   ScaleWidth      =   21831.14
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Customer entry.frx":65F1F
      Height          =   2295
      Left            =   3960
      TabIndex        =   19
      Top             =   7440
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   4048
      _Version        =   393216
      AllowArrows     =   0   'False
      Enabled         =   0   'False
      HeadLines       =   1
      RowHeight       =   15
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
            LCID            =   1033
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
            LCID            =   1033
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
   Begin VB.Frame Frame1 
      Caption         =   "Customer Entry"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   4680
      TabIndex        =   4
      Top             =   2640
      Width           =   9135
      Begin VB.CommandButton Command5 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   20
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2760
         TabIndex        =   13
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   855
         Left            =   2760
         TabIndex        =   11
         Top             =   1920
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2760
         TabIndex        =   10
         Top             =   3120
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   2760
         TabIndex        =   9
         Top             =   3600
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   6
         Top             =   2880
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Aadhar No.:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Pin code:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile no.:"
         BeginProperty Font 
            Name            =   "Sitka Text"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   3600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFFF00&
      Caption         =   "X"
      Height          =   195
      Left            =   17280
      TabIndex        =   3
      Top             =   3480
      Width           =   255
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Search"
      Height          =   2415
      Left            =   14400
      TabIndex        =   0
      Top             =   3480
      Width           =   3135
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
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   480
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   600
         Width           =   2175
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   840
      Top             =   6120
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=False"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from custdetail"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   600
      Picture         =   "Customer entry.frx":65F34
      Stretch         =   -1  'True
      Top             =   120
      Width           =   22260
   End
End
Attribute VB_Name = "customerdetail"
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
r.Source = " select * from custdetail"
r.CursorLocation = adUseClient
r.CursorType = adOpenDynamic
r.LockType = adLockBatchOptimistic
r.Open
Set customerdetail.DataGrid1.DataSource = r
customerdetail.DataGrid1.ReBind
End Function

Private Sub Command1_Click()
X
sql = " insert into custdetail values (" + Text1.Text + ",'" + Text2.Text + "','" + Text3.Text + "'," + Text4.Text + "," + Text5.Text + ")"

On Error GoTo handleError
Set r = c.Execute(sql)
Combo1.Clear

sql = "select adno from custdetail"
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
Text5.Text = ""
Text1.SetFocus
y
Exit Sub
handleError:
MsgBox " aadharnumber allready inserted", vbExclamation

Exit Sub


End Sub

Private Sub Command2_Click()
s = MsgBox(" are you want to delete this record", vbExclamation + vbYesNo)
 If s = vbYes Then
 X
sql = "delete from custdetail where adno = '" + Text1.Text + "' "
c.Execute (sql)
MsgBox "record deleted ", vbExclamation
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
y
Command2.Enabled = False
Command3.Enabled = False
Combo1.Refresh
End If
End Sub

Private Sub Command4_Click()
Frame2.Visible = True
Command8.Visible = True
End Sub

Private Sub Command3_Click()
sql = " update  custdetail set  nm = '" + Text2.Text + "', ads = '" + Text3.Text + "',pin = " + Text4.Text + ",mbno = " + Text5.Text + " where adno = " + Text1.Text + ""
Set r = c.Execute(sql)
MsgBox "record updated", vbInformation
y
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command5_Click()
Unload Me
Load customersale
customersale.Show
End Sub

Private Sub Command7_Click()
X
sql = "select * from custdetail where adno = '" + Combo1.Text + "'"

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
Text5.Text = r.Fields(4)
End If
DataGrid1.Enabled = False

MsgBox "record found", vbExclamation
Command2.Enabled = True
Command3.Enabled = True
Command8.Visible = False
Frame2.Visible = False
End If
End Sub

Private Sub Form_Load()
customerdetail.Height = 520
customerdetail.Width = 520
customerdetail.ScaleHeight = 520
customerdetail.ScaleWidth = 520
Combo1.Clear
Command2.Enabled = False
Command3.Enabled = False
Frame2.Visible = False
X

Command8.Visible = False
sql = "select adno from custdetail"
Set r = c.Execute(sql)
Do While (r.EOF <> True)
Combo1.AddItem r.Fields(0)
r.MoveNext
Loop
Adodc1.Visible = False
End Sub
Private Sub Command8_Click()
Frame2.Visible = False
Command8.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 46 Or KeyAscii = 13 Or KeyAscii = 8) Then
KeyAscii = KeyAscii
 If KeyAscii = 13 Then
 KeyAscii = KeyAscii
 Text2.SetFocus
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4.SetFocus
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
 End If
Else
MsgBox "please enter decimal values", vbExclamation
KeyAscii = 0
End If
End Sub
