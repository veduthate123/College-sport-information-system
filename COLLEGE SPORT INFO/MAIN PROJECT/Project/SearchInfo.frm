VERSION 5.00
Begin VB.Form frmSearchInfo 
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "SearchInfo"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form9"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   32
      Top             =   1680
      Width           =   20175
      Begin VB.CommandButton cmdhome 
         BackColor       =   &H00E0E0E0&
         Caption         =   "HOME"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   480
         Width           =   2295
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
         TabIndex        =   35
         Top             =   480
         Width           =   2175
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
         TabIndex        =   34
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdabout 
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
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   480
         Width           =   2295
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
      Height          =   1455
      Left            =   120
      TabIndex        =   18
      Top             =   240
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
         TabIndex        =   31
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         Left            =   9240
         TabIndex        =   20
         Top             =   840
         Width           =   1335
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
         Left            =   6120
         TabIndex        =   19
         Top             =   120
         Width           =   8175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   20175
      Begin VB.ComboBox Combo1 
         Height          =   405
         Left            =   8160
         TabIndex        =   1
         Text            =   "Search by Name"
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Search"
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
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label LblRM 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   12600
         TabIndex        =   30
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label LblADD 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10560
         TabIndex        =   29
         Top             =   6240
         Width           =   3975
      End
      Begin VB.Label LblEM 
         BackStyle       =   0  'Transparent
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
         Left            =   10560
         TabIndex        =   28
         Top             =   5160
         Width           =   3015
      End
      Begin VB.Label LblPP 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   27
         Top             =   5880
         Width           =   3015
      End
      Begin VB.Label LblPR 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   26
         Top             =   6840
         Width           =   3015
      End
      Begin VB.Label LblDOB 
         BackStyle       =   0  'Transparent
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
         Left            =   10560
         TabIndex        =   25
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label LblA 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   24
         Top             =   4080
         Width           =   3015
      End
      Begin VB.Label LblG 
         BackStyle       =   0  'Transparent
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
         Left            =   10560
         TabIndex        =   23
         Top             =   4200
         Width           =   3015
      End
      Begin VB.Label LblS 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   22
         Top             =   5040
         Width           =   3015
      End
      Begin VB.Label LblC 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   21
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Label LblLN 
         BackStyle       =   0  'Transparent
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
         Left            =   10560
         TabIndex        =   17
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label LblFN 
         BackStyle       =   0  'Transparent
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
         Left            =   4440
         TabIndex        =   16
         Top             =   2160
         Width           =   3015
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
         Left            =   14640
         TabIndex        =   15
         Top             =   2040
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   14640
         Stretch         =   -1  'True
         Top             =   2400
         Width           =   4095
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
         Left            =   8520
         TabIndex        =   14
         Top             =   6240
         Width           =   1335
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
         Left            =   2520
         TabIndex        =   13
         Top             =   6840
         Width           =   1455
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
         Left            =   2520
         TabIndex        =   12
         Top             =   5880
         Width           =   1455
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
         Left            =   8520
         TabIndex        =   11
         Top             =   5160
         Width           =   1095
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
         Left            =   2520
         TabIndex        =   10
         Top             =   5040
         Width           =   975
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
         Left            =   8520
         TabIndex        =   9
         Top             =   4200
         Width           =   1215
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
         Left            =   2520
         TabIndex        =   8
         Top             =   4080
         Width           =   735
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
         Left            =   8520
         TabIndex        =   7
         Top             =   3120
         Width           =   1935
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
         Left            =   2520
         TabIndex        =   6
         Top             =   3120
         Width           =   975
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
         Left            =   8520
         TabIndex        =   5
         Top             =   2040
         Width           =   1575
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
         Left            =   2520
         TabIndex        =   4
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Name :-"
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
         Left            =   5880
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmSearchInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim sqlQuery, vID, vFirstName, vLastName, vClass, vDateOfBirth, vAge, vGender, vSport, vEmail, vPhoneNo, vParentNo, vAddress As String
Dim vImgdata() As Byte

Private Sub cmdBack_Click()
Combo1.SetFocus
Hide
frmAdminPage.Show

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

Private Sub cmdabout_Click()
Combo1.SetFocus
     Combo1.Text = ""
     LblRM.Caption = ""
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
frmAbout.Show
frmSearchInfo.Hide

End Sub

Private Sub cmdADMINLOGIN_Click()
Combo1.SetFocus
     Combo1.Text = ""
     LblRM.Caption = ""
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
frmLogin.Show
frmSearchInfo.Hide
End Sub

Private Sub cmdContact_Click()
Combo1.SetFocus
     Combo1.Text = ""
     LblRM.Caption = ""
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
frmContact.Show
frmSearchInfo.Hide
End Sub

Private Sub cmdEnter_Click()
Dim vFirstName As String
vFirstName = Combo1.Text
sqlQuery = "SELECT * FROM tblname WHERE FirstName='" & vFirstName & "';"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    
    If ors.EOF = True Then
    
    LblRM.Caption = "Record Not Found"
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
     LblRM.Caption = ""
     
     
     
     End If
      conn.Close

Combo1.SetFocus
End Sub

Private Sub cmdHome_Click()
     Combo1.SetFocus
     Combo1.Text = ""
     LblRM.Caption = ""
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

frmHome.Show
frmSearchInfo.Hide
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()

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

