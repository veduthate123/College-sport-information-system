VERSION 5.00
Begin VB.Form frmReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame7 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   30
      Top             =   10320
      Width           =   12015
      Begin VB.Label Label1 
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
         Left            =   2520
         TabIndex        =   32
         Top             =   120
         Width           =   8535
      End
      Begin VB.Label Label17 
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
         Left            =   3840
         TabIndex        =   31
         Top             =   600
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8895
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   12015
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   8160
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5520
         TabIndex        =   1
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter First  Name :-"
         BeginProperty Font 
            Name            =   "Mongolian Baiti"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   29
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   28
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   27
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Class :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   26
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   23
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sport :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   4800
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   21
         Top             =   5400
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "PhoneNo :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "ParentNo :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   6600
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Address :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   3015
         Left            =   8280
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Image :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   17
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LblFN 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   16
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label LblLN 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Label LblC 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   14
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label LblS 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Top             =   4800
         Width           =   2895
      End
      Begin VB.Label LblG 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label LblA 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label LblDOB 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label LblPR 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   9
         Top             =   6600
         Width           =   2895
      End
      Begin VB.Label LblPP 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   8
         Top             =   6000
         Width           =   2895
      End
      Begin VB.Label LblEM 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   7
         Top             =   5400
         Width           =   2895
      End
      Begin VB.Label LblADD 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4920
         TabIndex        =   6
         Top             =   7200
         Width           =   4335
      End
   End
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
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.CommandButton Command1 
         Caption         =   "X"
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
         Left            =   11400
         TabIndex        =   34
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label15 
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
         Left            =   1920
         TabIndex        =   4
         Top             =   120
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim checkValid As String
Dim sqlQuery, vID, vFirstName, vLastName, vClass, vDateOfBirth, vAge, vGender, vSport, vEmail, vPhoneNo, vParentNo, vAddress As String
Dim vImgdata() As Byte

Private Sub cmdBack_Click()
Combo1.SetFocus
frmAdminPage.Show
frmReport.Hide

     
     Combo1.Text = ""
     LblFN.Caption = ""
     LblLN.Caption = ""
     LblC.Caption = ""
     LblDOB.Caption = ""
     LblA.Caption = ""
     LblG.Caption = ""
     LblS.Caption = ""
     LblEM.Caption = ""
     LblPP.Caption = ""
     LblPR.Caption = ""
     LblADD.Caption = ""
     Image1.Picture = Nothing
End Sub

Private Sub cmdEnter_Click()
checkValid = 0

If Combo1.Text = "" Then
MsgBox "Enter Name To Print Record ", vbExclamation, "Print"
Combo1.SetFocus
checkValid = checkValid + 1
End If



Dim vFirstName As String
vFirstName = Combo1.Text
sqlQuery = "SELECT * FROM tblname WHERE FirstName='" & vFirstName & "';"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    
    If ors.EOF = True Then
    
     LblFN.Caption = ""
     LblLN.Caption = ""
     LblC.Caption = ""
     LblDOB.Caption = ""
     LblA.Caption = ""
     LblG.Caption = ""
     LblS.Caption = ""
     LblEM.Caption = ""
     LblPP.Caption = ""
     LblPR.Caption = ""
     LblADD.Caption = ""
     Combo1.SetFocus
     
     

    Else
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
     vImgdata = ors.Fields("Img")
    
     
     LblFN.Caption = vFirstName
     LblLN.Caption = vLastName
     LblC.Caption = vClass
     LblDOB.Caption = vDateOfBirth
     LblA.Caption = vAge
     LblG.Caption = vGender
     LblS.Caption = vSport
     LblEM.Caption = vEmail
     LblPP.Caption = vPhoneNo
     LblPR.Caption = vParentNo
     LblADD.Caption = vAddress
     Image1.Picture = LoadPicture(vImgdata)
     
    
    End If
conn.Close


End Sub

Private Sub cmdPrint_Click()
PrintForm
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"


sqlQuery = "SELECT * FROM tblname;"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vFirstName = ors.Fields("FirstName").Value
      
     Combo1.AddItem vFirstName
     ors.MoveNext
    Loop
  
conn.Close

End Sub
