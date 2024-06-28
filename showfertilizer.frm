VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form showfertilizer 
   Caption         =   "Form1"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "showfertilizer.frx":0000
   ScaleHeight     =   11055
   ScaleMode       =   0  'User
   ScaleWidth      =   90771.58
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
      Connect         =   "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=True"
      OLEDBString     =   "Provider=MSDAORA.1;User ID=project/sonu;Persist Security Info=True"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from fertdetail"
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16560
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "showfertilizer.frx":65F1F
      Height          =   10095
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   17806
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   12648447
      ForeColor       =   16744448
      HeadLines       =   1
      RowHeight       =   21
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Fertilizer"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "FERTNM"
         Caption         =   "Fertilizer name"
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
         DataField       =   "FERTCOD"
         Caption         =   "Fertilizer code"
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
         Caption         =   "Company name"
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
         Caption         =   "MFG date"
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
         Caption         =   "Quantity"
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
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   9157.896
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   5805
      Left            =   15480
      Picture         =   "showfertilizer.frx":65F34
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "showfertilizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload showfertilizer
Load Fertilizer
Fertilizer.Top = 50
Fertilizer.Left = 10
Fertilizer.Show
Fertilizer.Top = 100
Fertilizer.Height = 11520
Fertilizer.Width = 20490
Fertilizer.ScaleHeight = 11055
Fertilizer.ScaleWidth = 20370
End Sub

Private Sub Picture1_Click()

End Sub
