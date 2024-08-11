VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAddinfo 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   " AddInfo"
   ClientHeight    =   11400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20250
   HasDC           =   0   'False
   LinkTopic       =   "Form8"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   10815
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
         Left            =   19560
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   615
      End
      Begin MSComDlg.CommonDialog CommDlg_Path 
         Left            =   19080
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11040
         TabIndex        =   4
         Top             =   3240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   124583939
         CurrentDate     =   44997
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   3960
         TabIndex        =   7
         Top             =   5400
         Width           =   2895
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
         Left            =   11040
         TabIndex        =   6
         Top             =   4320
         Width           =   2895
      End
      Begin VB.CommandButton cmd_LoadPicture 
         Caption         =   "Load Image"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   12
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   11040
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   6240
         Width           =   4215
      End
      Begin VB.TextBox txtParentNo 
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
         Left            =   3960
         TabIndex        =   10
         Top             =   7560
         Width           =   2895
      End
      Begin VB.TextBox txtPhoneNo 
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
         Left            =   3960
         TabIndex        =   9
         Top             =   6480
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   11040
         TabIndex        =   8
         Top             =   5280
         Width           =   2895
      End
      Begin VB.TextBox txtAge 
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
         Left            =   3960
         TabIndex        =   5
         Top             =   4320
         Width           =   2895
      End
      Begin VB.TextBox txtClass 
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
         Left            =   3960
         TabIndex        =   3
         Top             =   3240
         Width           =   2895
      End
      Begin VB.TextBox txtLastName 
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
         Left            =   11040
         TabIndex        =   2
         Top             =   2160
         Width           =   2895
      End
      Begin VB.TextBox txtFirstName 
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
         Left            =   3960
         TabIndex        =   1
         Top             =   2160
         Width           =   2895
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
         Left            =   10080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   9000
         Width           =   2415
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
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   9000
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Of Date :-"
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
         Left            =   9000
         TabIndex        =   27
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Image :-"
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
         Left            =   15480
         TabIndex        =   26
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Image Img 
         Height          =   3255
         Left            =   15480
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   4095
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Class:-"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label11 
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
         Left            =   9000
         TabIndex        =   24
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents No :-"
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
         Left            =   1680
         TabIndex        =   23
         Top             =   7560
         Width           =   1695
      End
      Begin VB.Label Label9 
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
         Height          =   255
         Left            =   9000
         TabIndex        =   22
         Top             =   6240
         Width           =   1335
      End
      Begin VB.Label Label8 
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
         Left            =   1680
         TabIndex        =   21
         Top             =   4320
         Width           =   735
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No :-"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   6480
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Sport Category :-"
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
         Left            =   1680
         TabIndex        =   19
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         Height          =   255
         Left            =   9000
         TabIndex        =   18
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Height          =   255
         Left            =   9000
         TabIndex        =   17
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   20160
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Player Registration"
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
         Left            =   7440
         TabIndex        =   15
         Top             =   600
         Width           =   4455
      End
   End
End
Attribute VB_Name = "frmAddinfo"
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
txtFirstName.SetFocus
frmAddinfo.Hide
frmAdminPage.Show

txtFirstName.Text = ""
txtLastName.Text = ""
txtClass.Text = ""
DTPicker1.Value = "01-Jan-2023"
txtAge.Text = ""
Combo1.Text = ""
Combo2.Text = ""
txtEmail.Text = ""
txtPhoneNo.Text = ""
txtParentNo.Text = ""
txtAddress.Text = ""
Img.Picture = LoadPicture("white.JPG")
End Sub

Private Sub cmdInsert_Click()

checkValid = 0
'Validation
If Len(txtFirstName.Text) = 0 Then
MsgBox "Enter First Name", , "First Name"
txtFirstName.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtLastName.Text) = 0 Then
MsgBox "Enter Last Name", , "Last Name"
txtLastName.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtClass.Text) = 0 Then
MsgBox "Enter Class", , "Class"
txtClass.SetFocus
checkValid = checkValid + 1
ElseIf DTPicker1.Value = 0 Then
MsgBox "Enter Birt Of Date", , "DOB"
DTPicker1.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtAge.Text) = 0 Then
MsgBox "Enter Age", , "Age"
txtAge.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtAge.Text) <= 1 Then
MsgBox "Age consist of 2 digit"
txtAge.SetFocus
txtAge = ""
checkValid = checkValid + 1
ElseIf Len(txtAge.Text) >= 3 Then
MsgBox "Age consist of 2 digit", , "Age"
txtAge.SetFocus
txtAge = ""
checkValid = checkValid + 1
ElseIf Combo1.Text = "" Then
MsgBox "Enter Gender", , "Gender"
Combo1.SetFocus
checkValid = checkValid + 1
ElseIf Combo2.Text = "" Then
MsgBox "Enter Sport", , "Sport"
Combo2.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtEmail.Text) = 0 Then
MsgBox "Enter Email", , "Email"
txtEmail.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtPhoneNo.Text) = 0 Then
MsgBox "Enter PhoneNo", , "Phone No"
txtPhoneNo.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtPhoneNo.Text) <= 9 Then
MsgBox "Phone No Must be 10 digit", , "Phonr No"
txtPhoneNo.SetFocus
txtPhoneNo = ""
checkValid = checkValid + 1
ElseIf Len(txtPhoneNo.Text) >= 11 Then
MsgBox "Phone No Must be 10 digit", , "Phonr No"
txtPhoneNo.SetFocus
txtPhoneNo = ""
checkValid = checkValid + 1
ElseIf Len(txtParentNo.Text) = 0 Then
MsgBox "Enter ParentNo", , "Parent No"
txtParentNo.SetFocus
checkValid = checkValid + 1
ElseIf Len(txtParentNo.Text) <= 9 Then
MsgBox "Parent No Must be 10 digit", , "Parent No"
txtParentNo.SetFocus
txtParentNo = ""
checkValid = checkValid + 1
ElseIf Len(txtParentNo.Text) >= 11 Then
MsgBox "Parent No Must be 10 digit", , "Parent No"
txtParentNo.SetFocus
txtParentNo = ""
checkValid = checkValid + 1
ElseIf Len(txtAddress.Text) = 0 Then
MsgBox "Enter Address", , "Address"
txtAddress.SetFocus
checkValid = checkValid + 1
End If


If checkValid = 0 Then
If MsgBox("Confirm add new student Record ?", vbQuestion + vbYesNo, "Collage Sport Information System") = vbYes Then
Dim ssql1 As String
ocon.Open
    ssql1 = ssql1 & "INSERT INTO tblname(FirstName,LastName,Class,BirthDate,Age,Gender,SportCategory,Email,PhoneNo,ParentNo,Address,Img)"
    ssql1 = ssql1 & " values("
    
    ssql1 = ssql1 & "'" & (txtFirstName.Text) & "',"
    ssql1 = ssql1 & "'" & (txtLastName.Text) & "',"
    ssql1 = ssql1 & "'" & (txtClass.Text) & "',"
    ssql1 = ssql1 & "'" & (DTPicker1.Value) & "',"
    ssql1 = ssql1 & "'" & (txtAge.Text) & "',"
    ssql1 = ssql1 & "'" & (Combo1.Text) & "',"
    ssql1 = ssql1 & "'" & (Combo2.Text) & "',"
    ssql1 = ssql1 & "'" & (txtEmail.Text) & "',"
    ssql1 = ssql1 & "'" & (txtPhoneNo.Text) & "',"
    ssql1 = ssql1 & "'" & (txtParentNo.Text) & "',"
    ssql1 = ssql1 & "'" & (txtAddress.Text) & "',"
    ssql1 = ssql1 & "'" & (CommDlg_Path.FileName) & "');"

ocon.Execute ssql1
ocon.Close
MsgBox "Player Registration successfully", vbInformation, "Collage Sport Information System"
txtFirstName.SetFocus
txtFirstName.Text = ""
txtLastName.Text = ""
txtClass.Text = ""
DTPicker1.Value = "01-Jan-2023"
txtAge.Text = ""
Combo1.Text = ""
Combo2.Text = ""
txtEmail.Text = ""
txtPhoneNo.Text = ""
txtParentNo.Text = ""
txtAddress.Text = ""
Img.Picture = LoadPicture("white.JPG")
End If
End If
End Sub

Private Sub Command1_Click()
End
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
TextDate.Text = Formate(DTPicker.Value, "custom")
End Sub

Private Sub Form_Load()
ocon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"

CommDlg_Path.FileName = App.Path & "\white.JPG"



Combo1.AddItem "Female"
Combo1.AddItem "Male"


Combo2.AddItem "Cricket"
Combo2.AddItem "Kabbadi"
Combo2.AddItem "VollyBall"
Combo2.AddItem "Badbinton"
Combo2.AddItem "TabelTannis"
Combo2.AddItem "Carrom"
Combo2.AddItem "Chess"
Combo2.AddItem "BasketBall"
Combo2.AddItem "FootBall"
Combo2.AddItem "Kho-Kho"

End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45) Then
KeyAscii = 0
End If
End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
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



