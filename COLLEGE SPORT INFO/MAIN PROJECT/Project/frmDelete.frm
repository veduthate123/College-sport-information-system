VERSION 5.00
Begin VB.Form frmDelete 
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
      TabIndex        =   4
      Top             =   1920
      Width           =   20175
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Modern No. 20"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   5
         Top             =   720
         Width           =   1935
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "Mongolian Baiti"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   13200
         TabIndex        =   2
         Top             =   720
         Width           =   2055
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
         Left            =   7320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   4935
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
         Left            =   8640
         TabIndex        =   7
         Top             =   240
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
         TabIndex        =   6
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
         TabIndex        =   3
         Top             =   240
         Width           =   8175
      End
   End
End
Attribute VB_Name = "frmDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim conn As New ADODB.Connection
Dim ors As New ADODB.Recordset
Dim checkValid As String
Dim sqlQuery, vFirstName As String

Private Sub cmdBack_Click()
Combo1.SetFocus
frmAdminPage.Show
frmDelete.Hide
End Sub

Private Sub cmdDelete_Click()
checkValid = 0

If Combo1.Text = "" Then
MsgBox "Enter FirstName to delete Record ", vbExclamation, "Delete"
Combo1.SetFocus
checkValid = checkValid + 1
End If

If checkValid = 0 Then
If MsgBox("Confirm Delete Student Record ?", vbQuestion + vbYesNo, "Collage Sport Information System") = vbYes Then
sqlQuery = "DELETE FROM tblname WHERE FirstName='" & Combo1.Text & "';"
conn.Open
conn.Execute sqlQuery
conn.Close
MsgBox "Record Deleted Successfully", vbInformation, "Collage Sport Information System"
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

End Sub

Private Sub Combo1_Change()

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

