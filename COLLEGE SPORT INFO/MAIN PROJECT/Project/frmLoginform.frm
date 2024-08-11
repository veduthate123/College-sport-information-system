VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   LinkTopic       =   "Form4"
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
      TabIndex        =   10
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
         TabIndex        =   12
         Top             =   840
         Width           =   5415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
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
         TabIndex        =   11
         Top             =   360
         Width           =   8535
      End
   End
   Begin VB.Frame Frame3 
      Height          =   9135
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   8175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Make A Dream...."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   9255
         Left            =   -1800
         Picture         =   "frmLoginform.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   8400
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      Begin VB.CheckBox Check1 
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton cmdHome 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Home"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdLOGIN 
         BackColor       =   &H00E0E0E0&
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5640
         Width           =   1575
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   4080
         Width           =   4215
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1800
         TabIndex        =   1
         Top             =   2400
         Width           =   4215
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name:-"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "LOGIN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery As String

Public LoginSucceeded As Boolean
Public vUsername As String
Public vPassword As String

Private Sub Check1_Click()
 If Check1.Value = 1 Then
 txtPassword.PasswordChar = ""
 ElseIf Check1.Value = 0 Then
 txtPassword.PasswordChar = "*"
 End If
 
End Sub

Private Sub cmdCancel_Click()
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Private Sub cmdBack_Click()
Form1.Show
Hide
End Sub

Private Sub cmdHome_Click()
txtUserName.Text = ""
txtPassword.Text = ""
txtUserName.SetFocus
frmHome.Show
frmLogin.Hide
End Sub

Private Sub cmdLOGIN_Click()

vUsername = txtUserName.Text
vPassword = txtPassword.Text

sqlQuery = "SELECT * FROM tbllogin WHERE username='" & vUsername & "' and password='" & vPassword & "';"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    If ors.EOF = True Then
     LoginSucceeded = False
    Else
     LoginSucceeded = True
     txtUserName.SetFocus
    End If
  
conn.Close

    If LoginSucceeded = True Then
     MsgBox "Login Successfully!", vbInformation, "Login"
        Me.Hide
        txtUserName = ""
        txtPassword = ""
       'Admin Form Show Here
        frmAdminPage.Show
        
    ElseIf Len(txtUserName.Text) = 0 Then
    MsgBox "Enter UserName", vbInformation, "User Name"
    txtUserName.SetFocus
    
    ElseIf Len(txtPassword.Text) = 0 Then
    MsgBox "Enter Password", vbInformation, "Password"
    txtPassword.SetFocus

    ElseIf LoginSucceeded = False Then
        MsgBox "Invalid Username Password, try again!", vbExclamation, "Login"
        txtUserName.SetFocus
        txtUserName = ""
        txtPassword = ""
    End If
    
End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\xyz.mdb;Persist Security Info=False"
End Sub


