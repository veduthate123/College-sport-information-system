VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddinfo 
   BackColor       =   &H80000007&
   Caption         =   " AddInfo"
   ClientHeight    =   7935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13320
   HasDC           =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   7935
   ScaleWidth      =   13320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin MSComDlg.CommonDialog CommDlg_Path 
         Left            =   12360
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmd_LoadPicture 
         Caption         =   "Load Image"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11280
         TabIndex        =   26
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   6960
         TabIndex        =   25
         Top             =   4440
         Width           =   3855
      End
      Begin VB.TextBox txtParentNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   24
         Top             =   4920
         Width           =   2535
      End
      Begin VB.TextBox txtPhoneNo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   23
         Top             =   4080
         Width           =   2535
      End
      Begin VB.TextBox txtEmail 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         TabIndex        =   22
         Top             =   3600
         Width           =   2535
      End
      Begin VB.ComboBox combo2 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmAddInfo.frx":0000
         Left            =   2280
         List            =   "frmAddInfo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3360
         Width           =   2535
      End
      Begin VB.ComboBox combo1 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmAddInfo.frx":0004
         Left            =   6960
         List            =   "frmAddInfo.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2760
         Width           =   2535
      End
      Begin VB.TextBox txtAge 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   19
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtBirthDate 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   18
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtClass 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   17
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtLastName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6960
         TabIndex        =   16
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   15
         Top             =   1200
         Width           =   2655
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Mongolian Baiti"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4560
         TabIndex        =   13
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton cmdInsert 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insert"
         BeginProperty Font 
            Name            =   "Mongolian Baiti"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Image :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10080
         TabIndex        =   27
         Top             =   1200
         Width           =   975
      End
      Begin VB.Image Img 
         Height          =   2175
         Left            =   10080
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Class:-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "E-Mail :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents No :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Address :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Age :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sport Category :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   5
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   4
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name :-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   13080
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Student Registration"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   1
         Top             =   120
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmAddInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ocon As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim checkValid As Integer
Dim ans

Private Sub cmd_LoadPicture_Click()
With CommDlg_Path
    .DialogTitle = "Search Employee picture"
    .Filter = "JPEG (Jpeg (*.jpg)|*.jpg|*.bmp)|*.bmp|Gif (*.gif)|*.gif|All Files (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    .ShowOpen
    .FilterIndex = 1
    .CancelError = False
Img.Picture = LoadPicture(.FileName)
End With
End Sub

Private Sub cmdCancel_Click()
frmAddInfo.Hide
frmAdminPage.Show
End Sub

Private Sub cmdInsert_Click()

checkValid = 0
'Validation
If Len(txtFirstName.Text) = 0 Then
MsgBox "Enter First Name"
txtFirstName.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtLastName.Text) = 0 Then
MsgBox "Enter Last Name"
txtLastName.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtClass.Text) = 0 Then
MsgBox "Enter Class"
txtClass.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtBirthDate.Text) = 0 Then
MsgBox "Enter BIrth Date"
txtBirthDate.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtAge.Text) = 0 Then
MsgBox "Enter Age"
txtAge.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtEmail.Text) = 0 Then
MsgBox "Enter Email"
txtEmail.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtPhoneNo.Text) = 0 Then
MsgBox "Enter PhoneNo"
txtPhoneNo.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtParentNo.Text) = 0 Then
MsgBox "Enter ParentNo"
txtParentNo.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtAddress.Text) = 0 Then
MsgBox "Enter Address"
txtAddress.SetFocus
checkValid = checkValid + 1
End If



If checkValid = 0 Then
If MsgBox("Confirm add new student Record !", vbQuestion + vbYesNo) = vbYes Then
Dim ssql1 As String
ocon.Open
    ssql1 = ssql1 & "INSERT INTO tblname(FirstName,LastName,Class,BirthDate,Age,Gender,SportCategory,Email,PhoneNo,ParentNo,Address,Img)"
    ssql1 = ssql1 & " values("
    
    ssql1 = ssql1 & "'" & (txtFirstName.Text) & "',"
    ssql1 = ssql1 & "'" & (txtLastName.Text) & "',"
    ssql1 = ssql1 & "'" & (txtClass.Text) & "',"
    ssql1 = ssql1 & "'" & (txtBirthDate.Text) & "',"
    ssql1 = ssql1 & "'" & (txtAge.Text) & "',"
    ssql1 = ssql1 & "'" & (combo1.Text) & "',"
    ssql1 = ssql1 & "'" & (combo2.Text) & "',"
    ssql1 = ssql1 & "'" & (txtEmail.Text) & "',"
    ssql1 = ssql1 & "'" & (txtPhoneNo.Text) & "',"
    ssql1 = ssql1 & "'" & (txtParentNo.Text) & "',"
    ssql1 = ssql1 & "'" & (txtAddress.Text) & "',"
    ssql1 = ssql1 & "'" & (CommDlg_Path.FileName) & "');"

ocon.Execute ssql1
ocon.Close
MsgBox "Recordset added successfully"

txtFirstName.Text = ""
txtLastName.Text = ""
txtClass.Text = ""
txtBirthDate.Text = ""
txtAge.Text = ""
txtEmail.Text = ""
txtPhoneNo.Text = ""
txtParentNo.Text = ""
txtAddress.Text = ""
Img.Picture = LoadPicture("white.JPG")
End If
End If
End Sub

Private Sub Form_Load()
ocon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"

CommDlg_Path.FileName = App.Path & "\white.JPG"



combo1.AddItem "Female"
combo1.AddItem "Male"


combo2.AddItem "Cricket"
combo2.AddItem "Kabbadi"
combo2.AddItem "VollyBall"
combo2.AddItem "Badbinton"
combo2.AddItem "TabelTannis"
combo2.AddItem "Carrom"
combo2.AddItem "Chess"
combo2.AddItem "BasketBall"
combo2.AddItem "FootBall"
combo2.AddItem "Kho-Kho"

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtBirthDate_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 92 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtClass_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 64 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 64 Or KeyAscii = 46 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtLastName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtParentNo_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtParentNo_Validate(Cancel As Boolean)
If Len(txtParentNo.Text) >= 10 Then
MsgBox "Parent No must be 10 Digit"
txtParentNo = ""
txtParentNo.SetFocus
checkValid = checkValid + 1
End If
End Sub

Private Sub txtPhoneNo_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtFirstName_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtPhoneNo_Validate(Cancel As Boolean)
If Len(txtPhoneNo.Text) >= 10 Then
MsgBox "Phone No must be 10 Digit"
txtPhoneNo = ""
checkValid = checkValid + 1
End If
txtPhoneNo.SetFocus
End Sub
