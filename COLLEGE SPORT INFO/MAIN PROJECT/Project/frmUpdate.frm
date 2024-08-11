VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpdate 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Update"
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
   LinkTopic       =   "Form5"
   ScaleHeight     =   11400
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   9255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   20175
      Begin MSComDlg.CommonDialog CommDlg_Path 
         Left            =   18240
         Top             =   5880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Load Picture"
         Height          =   615
         Left            =   16440
         TabIndex        =   16
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox txtADD 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10920
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   6360
         Width           =   3855
      End
      Begin VB.TextBox txtPR 
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
         Left            =   4440
         TabIndex        =   14
         Top             =   7320
         Width           =   3015
      End
      Begin VB.TextBox txtPP 
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
         Left            =   4440
         TabIndex        =   13
         Top             =   6240
         Width           =   3015
      End
      Begin VB.TextBox txtEM 
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
         Left            =   10920
         TabIndex        =   12
         Top             =   5280
         Width           =   3015
      End
      Begin VB.TextBox txtS 
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
         Left            =   4440
         TabIndex        =   11
         Top             =   5160
         Width           =   3015
      End
      Begin VB.TextBox txtG 
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
         Left            =   10920
         TabIndex        =   10
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox txtA 
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
         Left            =   4440
         TabIndex        =   9
         Top             =   4200
         Width           =   3015
      End
      Begin VB.TextBox txtDOB 
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
         Left            =   10920
         TabIndex        =   8
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtC 
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
         Left            =   4440
         TabIndex        =   7
         Top             =   3240
         Width           =   3015
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
         Left            =   8760
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
      Begin VB.CommandButton cmdEnter 
         Caption         =   "Enter"
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
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
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
         Left            =   9240
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
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
         Left            =   12240
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtFN 
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
         Left            =   4440
         TabIndex        =   5
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtLN 
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
         Left            =   10920
         TabIndex        =   6
         Top             =   2160
         Width           =   3015
      End
      Begin VB.Image Image1 
         Height          =   3495
         Left            =   15240
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   3975
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
         Left            =   15240
         TabIndex        =   31
         Top             =   1920
         Width           =   975
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
         Height          =   255
         Left            =   8760
         TabIndex        =   30
         Top             =   6480
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
         Left            =   2040
         TabIndex        =   29
         Top             =   7320
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
         Left            =   2040
         TabIndex        =   28
         Top             =   6360
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
         Left            =   8760
         TabIndex        =   27
         Top             =   5400
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
         Left            =   2040
         TabIndex        =   26
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter First  Name :-"
         BeginProperty Font 
            Name            =   "Mongolian Baiti"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6120
         TabIndex        =   25
         Top             =   360
         Width           =   2535
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
         Left            =   2040
         TabIndex        =   24
         Top             =   2160
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
         Left            =   8760
         TabIndex        =   23
         Top             =   2160
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
         Left            =   2040
         TabIndex        =   22
         Top             =   3240
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
         Left            =   8760
         TabIndex        =   21
         Top             =   3240
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
         Left            =   2040
         TabIndex        =   20
         Top             =   4200
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
         Left            =   8760
         TabIndex        =   19
         Top             =   4200
         Width           =   1215
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
      TabIndex        =   0
      Top             =   360
      Width           =   20175
      Begin VB.CommandButton Command2 
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
         TabIndex        =   32
         Top             =   120
         Width           =   615
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
         TabIndex        =   17
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmUpdate"
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
Hide
frmAdminPage.Show

     Combo1.Text = ""
     txtFN.Text = ""
     txtLN.Text = ""
     txtC.Text = ""
     txtDOB.Text = ""
     txtA.Text = ""
     txtG.Text = ""
     txtS.Text = ""
     txtEM.Text = ""
     txtPP.Text = ""
     txtPR.Text = ""
     txtADD.Text = ""
     Image1.Picture = Nothing
End Sub

Private Sub cmdEnter_Click()
Dim vFirstName As String
vFirstName = Combo1.Text
sqlQuery = "SELECT * FROM tblname WHERE FirstName='" & vFirstName & "';"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    
    If ors.EOF = True Then

     txtFN.Text = "Record Not Found"
     txtLN.Text = "Record Not Found"
     txtC.Text = "Record Not Found"
     txtDOB.Text = "Record Not Found"
     txtA.Text = "Record Not Found"
     txtG.Text = "Record Not Found"
     txtS.Text = "Record Not Found"
     txtEM.Text = "Record Not Found"
     txtPP.Text = "Record Not Found"
     txtPR.Text = "Record Not Found"
     txtADD.Text = "Record Not Found"
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
    
     
     txtFN.Text = vFirstName
     txtLN.Text = vLastName
     txtC.Text = vClass
     txtDOB.Text = vDateOfBirth
     txtA.Text = vAge
     txtG.Text = vGender
     txtS.Text = vSport
     txtEM.Text = vEmail
     txtPP.Text = vPhoneNo
     txtPR.Text = vParentNo
     txtADD.Text = vAddress
     Image1.Picture = LoadPicture(vImgdata)
     
     
    End If
  
conn.Close

Combo1.SetFocus
End Sub

Private Sub cmdUpdate_Click()
checkValid = 0



    If Len(txtFN.Text) = 0 Then
        MsgBox "First name cannot be empty", , "Firstname"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtLN.Text) = 0 Then
        MsgBox "Last Name cannot be empty", , "Lastname"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtC.Text) = 0 Then
        MsgBox "Class cannot be empty", , "Class"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtDOB.Text) = 0 Then
        MsgBox "Date of Birth cannot be empty", , "BirthDate"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtA.Text) = 0 Then
        MsgBox "Age not be empty", , "Age"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtG.Text) = 0 Then
        MsgBox "Gender not be empty", , "Gender"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtS.Text) = 0 Then
        MsgBox "Sport not be empty", , "Sport"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtEM.Text) = 0 Then
        MsgBox "Email not be empty", , "Email"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtPP.Text) = 0 Then
        MsgBox "Phone No not be empty", , "PhoneNo"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtPR.Text) = 0 Then
        MsgBox "Parent No not be empty", , "ParentNo"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If Len(txtADD.Text) = 0 Then
        MsgBox "Address not be empty", , "Address"
        Combo1.SetFocus
        checkValid = checkValid + 1
    End If
    
    If checkValid > 0 Then
        Exit Sub ' Exit the sub if validation fails
    End If

If Combo1.Text = "" Then
MsgBox "Search Record First ", vbExclamation, "Update"
Combo1.SetFocus
checkValid = checkValid + 1
End If

If checkValid = 0 Then
If MsgBox("Sure want to Update Record", vbQuestion + vbYesNo, "Collage Sport Information System") = vbYes Then

Dim vFirstName As String
vFirstName = Combo1.Text
sqlQuery = "SELECT * FROM tblname WHERE FirstName='" & vFirstName & "';"
conn.Open
    ors.Open sqlQuery, conn, adOpenDynamic, adLockOptimistic
    If Not ors.EOF Then
    ors.Update "FirstName", txtFN.Text
    ors.Update "LastName", txtLN.Text
    ors.Update "Class", txtC.Text
    ors.Update "BirthDate", txtDOB.Text
    ors.Update "Age", txtA.Text
    ors.Update "Gender", txtG.Text
    ors.Update "SportCategory", txtS.Text
    ors.Update "Email", txtEM.Text
    ors.Update "PhoneNo", txtPP.Text
    ors.Update "ParentNo", txtPR.Text
    ors.Update "Address", txtADD.Text
    
    
    With ors
        If Image1 <> "" Then
             .Fields("Img").AppendChunk CommDlg_Path.FileName
             .Update
          'load the new image into the Image1 control
             If CommDlg_Path.FileName <> "" Then
             Image1.Picture = LoadPicture(CommDlg_Path.FileName)
             End If
             End If
    End With
    
     
    MsgBox "Record Updated Successfully", vbInformation, "Collage Sport Information System"
    conn.Close


    End If
    End If
    End If
    
Combo1.Clear

sqlQuery = "SELECT * FROM tblname;"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vFirstName = ors.Fields("FirstName").Value
      
     Combo1.AddItem vFirstName
     ors.MoveNext
    Loop
  
conn.Close

   Combo1.SetFocus
  Combo1.Text = ""
     txtFN.Text = ""
     txtLN.Text = ""
     txtC.Text = ""
     txtDOB.Text = ""
     txtA.Text = ""
     txtG.Text = ""
     txtS.Text = ""
     txtEM.Text = ""
     txtPP.Text = ""
     txtPR.Text = ""
     txtADD.Text = ""
     Image1.Picture = Nothing

End Sub

Private Sub Command1_Click()
With CommDlg_Path
    .DialogTitle = "Search Employee picture"
    .Filter = "JPEG (Jpeg (*.jpg)|*.jpg|*.bmp)|*.bmp|Gif (*.gif)|*.gif|All Files (*.*)|*.*"
    .Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt + cdlOFNPathMustExist
    .ShowOpen
    .FilterIndex = 1
    .CancelError = False
Image1.Picture = LoadPicture(.FileName)
End With
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()

conn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "\database\rd.mdb;Persist Security Info=False"
CommDlg_Path.FileName = App.Path & "\white.JPG"

sqlQuery = "SELECT * FROM tblname;"

conn.Open
    ors.Open sqlQuery, conn, adOpenForwardOnly, adLockReadOnly
    Do Until ors.EOF
     vFirstName = ors.Fields("FirstName").Value
      
     Combo1.AddItem vFirstName
     ors.MoveNext
    Loop
  
conn.Close

Dim imagePath As String
    
    ' Set the relative path to the image file
    imagePath = "Images\picture.jpg"
    
    ' Check if the image file exists
    If Dir(imagePath) <> "" Then
        ' Load the image into the PictureBox control
        On Error Resume Next ' temporarily ignore errors for the next line
        Set Image1.Picture = LoadPicture(imagePath)
        On Error GoTo 0 ' re-enable error checking
        If Err.Number <> 0 Then
            MsgBox "Error loading image: " & Err.Description
        End If
    Else
        
    End If
End Sub



Private Sub txtA_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If

End Sub

Private Sub txtADD_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46 Or KeyAscii = 45) Then
KeyAscii = 0
End If

End Sub

Private Sub txtC_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 32 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 45 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtEM_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 64 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 64 Or KeyAscii = 46 Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtFN_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtG_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If

End Sub

Private Sub txtLN_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtPP_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtPR_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub

Private Sub txtS_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8) Then
KeyAscii = 0
End If
End Sub
