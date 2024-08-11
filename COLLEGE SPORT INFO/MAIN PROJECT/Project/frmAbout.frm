VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "About"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   FillColor       =   &H00FFC0C0&
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form2"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   11
      Top             =   9600
      Width           =   20175
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
         TabIndex        =   13
         Top             =   360
         Width           =   8535
      End
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
         TabIndex        =   12
         Top             =   840
         Width           =   5415
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   1800
      Width           =   20175
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdContact 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CONTACT"
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
         Left            =   16800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   480
         Width           =   2295
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
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
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
      TabIndex        =   7
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
         TabIndex        =   14
         Top             =   120
         Width           =   615
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
         TabIndex        =   9
         Top             =   120
         Width           =   8175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "About Us"
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
         Left            =   9000
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   9960
      TabIndex        =   5
      Top             =   3600
      Width           =   10335
      Begin VB.Label Label4 
         Caption         =   $"frmAbout.frx":0000
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   9735
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   9735
      Begin VB.Image Image1 
         Height          =   6120
         Left            =   0
         Picture         =   "frmAbout.frx":02C0
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   9780
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdADMINLOGIN_Click()
frmAbout.Hide
frmLogin.Show
End Sub

Private Sub cmdContact_Click()
frmContact.Show
frmAbout.Hide
End Sub

Private Sub cmdHome_Click()
frmHome.Show
frmAbout.Hide
End Sub

Private Sub cmdSEARCH_Click()
frmSearchInfo.Show
frmAbout.Hide
End Sub

Private Sub Command1_Click()
End
End Sub
