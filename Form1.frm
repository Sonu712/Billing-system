VERSION 5.00
Begin VB.Form login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Login"
   ClientHeight    =   6510
   ClientLeft      =   4080
   ClientTop       =   1830
   ClientWidth     =   13020
   DrawStyle       =   1  'Dash
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "LOGIN"
      Height          =   735
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Lucida Bright"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3120
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   5
         ToolTipText     =   "Enter your Paasword"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Please Enter user Name"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget Paasword?"
         BeginProperty Font 
            Name            =   "Rockwell"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         ToolTipText     =   "Can't acsess Account?"
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1920
      Left            =   960
      Picture         =   "Form1.frx":21DA3
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To JMFS Managenent"
      BeginProperty Font 
         Name            =   "Sitka Text"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   975
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   6060
      Left            =   120
      MousePointer    =   1  'Arrow
      Picture         =   "Form1.frx":2B0B1
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12780
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Label4_Click()
Load Form2
Form2.Show
End Sub
