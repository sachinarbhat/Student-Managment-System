VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18120
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Student.frx":0000
   ScaleHeight     =   9135
   ScaleWidth      =   18120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "QUERRY"
      Height          =   375
      Left            =   15240
      TabIndex        =   23
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "REPORT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12720
      TabIndex        =   22
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   13440
      TabIndex        =   21
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   11520
      TabIndex        =   20
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      DataField       =   "DOB"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   2760
      Width           =   2895
   End
   Begin VB.OptionButton Option2 
      Caption         =   "F"
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   4200
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "M"
      Height          =   375
      Left            =   7200
      TabIndex        =   17
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MAIN MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   16
      Top             =   6720
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7440
      TabIndex        =   15
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "REMOVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      Top             =   6720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   6720
      Width           =   2055
   End
   Begin VB.ComboBox Combo3 
      DataField       =   "CID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      TabIndex        =   12
      Text            =   "select course"
      Top             =   5640
      Width           =   2895
   End
   Begin VB.ComboBox Combo2 
      DataField       =   "DID"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      TabIndex        =   11
      Text            =   "select departmeent"
      Top             =   4920
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      DataField       =   "G mail"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   3480
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   9
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "USN"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   8
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   7
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "GENDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "GMAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF BIRTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "STUDENT PROFILE"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   9015
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo2_Click()
Combo3.Clear
If Combo2.Text = "CSE" Then
Combo3.AddItem "jav"
Combo3.AddItem "uni"
Combo3.AddItem "pyt"
ElseIf Combo2.Text = "EEE" Then
Combo3.AddItem "em"
Combo3.AddItem "cs"
Combo3.AddItem "cae"
ElseIf Combo2.Text = "ME" Then
Combo3.AddItem "mom"
Combo3.AddItem "dom"
Combo3.AddItem "cad"
ElseIf Combo2.Text = "CVE" Then
Combo3.AddItem "som"
Combo3.AddItem "am"
Combo3.AddItem "sur"
Else
 End If
 
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command7_Click()
Form4.Hide
DataReport3.Show

End Sub

Private Sub Command8_Click()
Form4.Hide
Form8.Show
End Sub

Private Sub Form_Load()
Combo2.AddItem "CSE"
Combo2.AddItem "EEE"
Combo2.AddItem "ME"
Combo2.AddItem "CVE"
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update


End Sub

Private Sub Command4_Click()
Form4.Hide
Form2.Show
End Sub


Private Sub Option1_Click()
Dim value As String
value = ""
isChecked = Option1.Enabled

If (isChecked) Then
value = "M"
Adodc1.Recordset.Fields("Gender") = "M"
Else
End If
End Sub

Private Sub Option2_Click()
Dim value As String
value = ""
isChecked = Option2.Enabled

If (isChecked) Then
value = "F"
Adodc1.Recordset.Fields("Gender") = "F"
Else
End If
End Sub

