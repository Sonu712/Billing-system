VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form showseed 
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "showseed.frx":0000
   ScaleHeight     =   11055
   ScaleMode       =   0  'User
   ScaleWidth      =   90771.58
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Close"
      Height          =   615
      Left            =   17280
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7080
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "showseed.frx":65F1F
      Height          =   10095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   17806
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12640511
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      Caption         =   "SEED"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "SEEDNM"
         Caption         =   "Seed Name"
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
         DataField       =   "SEEDCOD"
         Caption         =   "Seed Code"
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
      BeginProperty Column02 
         DataField       =   "COMPNYNM"
         Caption         =   "Company Name"
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
      BeginProperty Column03 
         DataField       =   "BCHNO"
         Caption         =   "Batch No."
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
      BeginProperty Column04 
         DataField       =   "MFGDT"
         Caption         =   "MFG Date"
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
      BeginProperty Column05 
         DataField       =   "EXPDT"
         Caption         =   "EXP Date"
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
      BeginProperty Column06 
         DataField       =   "QTYBAG"
         Caption         =   "Quantity/Bag"
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
      BeginProperty Column07 
         DataField       =   "PCEBAG"
         Caption         =   "Price/Bag"
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
      BeginProperty Column08 
         DataField       =   "PCEKG"
         Caption         =   "Price/Kg"
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
            Alignment       =   2
            ColumnWidth     =   8756.211
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            ColumnWidth     =   7485.474
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   11229.47
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   6818.527
         EndProperty
         BeginProperty Column06 
            Alignment       =   2
            ColumnWidth     =   8890.105
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   1215
      Left            =   4800
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   2143
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
      RecordSource    =   "select * from seeddetail "
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
      Height          =   5805
      Left            =   16800
      Picture         =   "showseed.frx":65F34
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   3030
   End
End
Attribute VB_Name = "showseed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload showseed
Load Seeds
Seeds.Show
Seeds.Height = 11520
Seeds.Width = 20490
Seeds.ScaleHeight = 11055
Seeds.ScaleWidth = 20370
End Sub


