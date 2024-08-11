VERSION 5.00
Begin VB.Form frmContact 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Contact"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   20175
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   19440
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Us"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   19
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "College Sport Information System"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   18
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   20175
      Begin VB.CommandButton cmdADMINLOGIN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ADMIN LOGIN"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.CommandButton cmdSEARCH 
         BackColor       =   &H00E0E0E0&
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   480
         Width           =   2250
      End
      Begin VB.CommandButton cmdABOUT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ABOUT"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdHOME 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   16320
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   13
      Top             =   9600
      Width           =   20175
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Guided By Asst. Prof. Atul Tayde"
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   15
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Design And Developed By Vedant Thate..."
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   20175
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "9922587415"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   17520
         TabIndex        =   12
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "1(234)567-891"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   14040
         TabIndex        =   11
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sport@Info.Com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   17400
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Lat's get the conversation started."
         BeginProperty Font 
            Name            =   "Ravie"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13680
         TabIndex        =   9
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "sportlive@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   13800
         TabIndex        =   8
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Contect  :- 9922587415"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   3480
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail :- Sportlive@gmail.com"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   6
         Top             =   2280
         Width           =   6855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address :- Collage Sport Department"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   7335
      End
      Begin VB.Image Image1 
         Height          =   6435
         Left            =   -240
         Picture         =   "frmContact.frx":0000
         Stretch         =   -1  'True
         Top             =   -480
         Width           =   20610
      End
   End
End
Attribute VB_Name = "frmContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdABOUT_Click()
frmContact.Hide
frmAbout.Show
End Sub

Private Sub cmdADMINLOGIN_Click()
frmContact.Hide
frmLogin.Show
End Sub

Private Sub cmdHome_Click()
frmHome.Show
frmContact.Hide
End Sub

Private Sub cmdSEARCH_Click()
frmSearchInfo.Show
frmContact.Hide
End Sub

Private Sub Command1_Click()
End
End Sub
