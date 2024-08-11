VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStarting 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "Starting"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   LinkTopic       =   "Form6"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   10935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   20160
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   360
         Top             =   600
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   7200
         Width           =   18135
         _ExtentX        =   31988
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Max             =   105
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   10200
         TabIndex        =   4
         Top             =   6840
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Cooper Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   7320
         TabIndex        =   2
         Top             =   6840
         Width           =   2775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To My Project"
         BeginProperty Font 
            Name            =   "Wide Latin"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   5040
         TabIndex        =   1
         Top             =   2640
         Width           =   9015
      End
   End
End
Attribute VB_Name = "frmStarting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 5
Label2.Caption = "Project Loading...."
Label3.Caption = ProgressBar1.Value & "%"
If (ProgressBar1.Value = ProgressBar1.Max) Then
MsgBox "Loading Successfull!", , ""
Timer1.Enabled = False
frmHome.Show
Hide
Unload Me
End If
End Sub
