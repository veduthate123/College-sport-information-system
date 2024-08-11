VERSION 5.00
Begin VB.Form frmShowData 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   20175
      Begin VB.CommandButton cmdShowData 
         Caption         =   "Show Data"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   11160
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   240
         ScaleHeight     =   7095
         ScaleWidth      =   735
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   1080
         ScaleHeight     =   7095
         ScaleWidth      =   1335
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   2520
         ScaleHeight     =   7095
         ScaleWidth      =   1335
         TabIndex        =   14
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   3960
         ScaleHeight     =   7095
         ScaleWidth      =   1095
         TabIndex        =   13
         Top             =   2040
         Width           =   1095
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   5160
         ScaleHeight     =   7095
         ScaleWidth      =   1335
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   6600
         ScaleHeight     =   7095
         ScaleWidth      =   855
         TabIndex        =   11
         Top             =   2040
         Width           =   855
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   7560
         ScaleHeight     =   7095
         ScaleWidth      =   1335
         TabIndex        =   10
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   9000
         ScaleHeight     =   7095
         ScaleWidth      =   1335
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   10440
         ScaleHeight     =   7095
         ScaleWidth      =   2535
         TabIndex        =   8
         Top             =   2040
         Width           =   2535
      End
      Begin VB.PictureBox Picture10 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   13080
         ScaleHeight     =   7095
         ScaleWidth      =   1455
         TabIndex        =   7
         Top             =   2040
         Width           =   1455
      End
      Begin VB.PictureBox Picture11 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   14640
         ScaleHeight     =   7095
         ScaleWidth      =   1455
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.PictureBox Picture12 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7095
         Left            =   16200
         ScaleHeight     =   7095
         ScaleWidth      =   3855
         TabIndex        =   5
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6840
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label13 
         Caption         =   "Phone No :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   13080
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Parent No :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14640
         TabIndex        =   28
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Address :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   16200
         TabIndex        =   27
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label10 
         Caption         =   "DOB :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Age :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Gender :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7560
         TabIndex        =   24
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "ID :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Firstname :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Class :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   21
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Lastname :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Sport :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9000
         TabIndex        =   19
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Email :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10440
         TabIndex        =   18
         Top             =   1800
         Width           =   2535
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
      TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery, vID, vFirstName, vLastName, vClass, vDateOfBirth, vAge, vGender, vSport, vEmail, vPhoneNo, vParentNo, vAddress As String


Private Sub cmdBack_Click()
frmAdminPage.Show
frmShowData.Hide
End Sub

Private Sub cmdShowData_Click()
sqlQuery = "SELECT * FROM tblname;"

 conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vID = ors.Fields("id").Value
     vFirstName = ors.Fields("FirstName").Value
     vLastName = ors.Fields("LastName").Value
     vClass = ors.Fields("Class").Value
     vDateOfBirth = ors.Fields("BirthDate").Value
     vAge = ors.Fields("Age").Value
     vGender = ors.Fields("Gender").Value
     vSport = ors.Fields("SportCategory").Value
     vEmail = ors.Fields("Email").Value
     vPhoneNo = ors.Fields("PhoneNo").Value
     vParentNo = ors.Fields("ParentNo").Value
     vAddress = ors.Fields("Address").Value
     
    Picture1.Print vID
    Picture2.Print vFirstName
    Picture3.Print vLastName
    Picture4.Print vClass
    Picture5.Print vDateOfBirth
    Picture6.Print vAge
    Picture7.Print vGender
    Picture8.Print vSport
    Picture9.Print vEmail
    Picture10.Print vPhoneNo
    Picture11.Print vParentNo
    Picture12.Print vAddress
     ors.MoveNext
     
    Loop
  
conn.Close
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"

conn.Close

End Sub




