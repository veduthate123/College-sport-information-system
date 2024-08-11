VERSION 5.00
Begin VB.Form frmAdminPage 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Admin"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20355
   LinkTopic       =   "Form7"
   ScaleHeight     =   11400
   ScaleWidth      =   20355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
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
      TabIndex        =   13
      Top             =   360
      Width           =   20175
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Page"
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
         Left            =   8640
         TabIndex        =   15
         Top             =   840
         Width           =   2175
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
         TabIndex        =   14
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   20175
      Begin VB.CommandButton cmdShow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data Show"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   7680
         Width           =   2055
      End
      Begin VB.CommandButton cmdDataShow 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data Show"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5160
         Width           =   2175
      End
      Begin VB.CommandButton cmdlog 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LogOut"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   8040
         Width           =   2265
      End
      Begin VB.CommandButton cmdRprint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9720
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   7680
         Width           =   2145
      End
      Begin VB.CommandButton cmdDeleteInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5880
         Width           =   2055
      End
      Begin VB.CommandButton cmdUpdateInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5880
         Width           =   2145
      End
      Begin VB.CommandButton cmdAddInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5880
         Width           =   2145
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Delete Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Update Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2040
         Width           =   2145
      End
      Begin VB.CommandButton cmdLogOut 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LogOut"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   14400
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   7680
         Width           =   2145
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   6600
         Width           =   2265
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Add Info"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   2145
      End
      Begin VB.Line Line3 
         X1              =   20040
         X2              =   20040
         Y1              =   5400
         Y2              =   -120
      End
      Begin VB.Line Line1 
         X1              =   3720
         X2              =   20400
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Line Line2 
         X1              =   3720
         X2              =   3720
         Y1              =   0
         Y2              =   10800
      End
      Begin VB.Image Image1 
         Height          =   5175
         Left            =   3840
         Picture         =   "frmAdminPage.frx":0000
         Stretch         =   -1  'True
         Top             =   120
         Width           =   16095
      End
   End
End
Attribute VB_Name = "frmAdminPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
frmAddinfo.Show
frmAdminPage.Hide
End Sub

Private Sub cmdAddInfo_Click()
frmAddinfo.Show
frmAdminPage.Hide
End Sub

Private Sub cmdDataShow_Click()
frmShowData.Show
frmAdminPage.Hide
End Sub

Private Sub cmdDelete_Click()
frmDelete.Show
frmAdminPage.Hide
End Sub

Private Sub cmdDeleteInfo_Click()
frmDelete.Show
frmAdminPage.Hide
End Sub

Private Sub cmdlog_Click()
If MsgBox("Confirm Logout ?", vbQuestion + vbYesNo, "LogOut") = vbYes Then
frmAdminPage.Hide
frmHome.Show
MsgBox "LogOut Successfully", , "LogOut"
End If
End Sub

Private Sub cmdLogOut_Click()
If MsgBox("Confirm Logout ?", vbQuestion + vbYesNo, "LogOut") = vbYes Then
frmAdminPage.Hide
frmHome.Show
MsgBox "LogOut Successfully", , "LogOut"
End If
End Sub

Private Sub cmdSEARCH_Click()
frmSearchInfo.Show
frmAdminPage.Hide
End Sub

Private Sub cmdReport_Click()
frmReport.Show
frmAdminPage.Hide
End Sub

Private Sub cmdRprint_Click()
frmReport.Show
frmAdminPage.Hide
End Sub

Private Sub cmdShow_Click()
frmShowData.Show
frmAdminPage.Hide
End Sub

Private Sub cmdUpdate_Click()
frmUpdate.Show
frmAdminPage.Hide
End Sub

Private Sub cmdUpdateInfo_Click()
frmUpdate.Show
frmAdminPage.Hide
End Sub


