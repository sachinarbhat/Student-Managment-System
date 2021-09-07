VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   LinkTopic       =   "Form6"
   Picture         =   "Department.frx":0000
   ScaleHeight     =   7920
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      DataField       =   "Dname"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   3600
      Width           =   3855
   End
   Begin VB.CommandButton Command6 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   14160
      TabIndex        =   9
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   12240
      TabIndex        =   8
      Top             =   5520
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      DataField       =   "DID"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   6840
      TabIndex        =   7
      Text            =   "select department id"
      Top             =   2040
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10200
      TabIndex        =   6
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7560
      TabIndex        =   5
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4920
      TabIndex        =   4
      Top             =   6120
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT NAME"
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
      Height          =   735
      Left            =   3960
      TabIndex        =   2
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTMENT"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   10935
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command5_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Form_Load()
Combo1.AddItem "CSE"
Combo1.AddItem "ME"
Combo1.AddItem "EEE"
Combo1.AddItem "CVE"
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
Form6.Hide
Form2.Show
End Sub

