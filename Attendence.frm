VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "AQS"
   ClientHeight    =   7545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15405
   LinkTopic       =   "Form3"
   Picture         =   "Attendence.frx":0000
   ScaleHeight     =   7545
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command8 
      Caption         =   "Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13440
      TabIndex        =   14
      Top             =   6480
      Width           =   1695
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   13080
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   10920
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   11
      Top             =   3600
      Width           =   735
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
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Top             =   6480
      Width           =   2055
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
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   6480
      Width           =   2055
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
      Height          =   495
      Left            =   6000
      TabIndex        =   8
      Top             =   6480
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
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   6480
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "percentage"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   3840
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "class_attended"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   3120
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      DataField       =   "USN"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   2280
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PERCENTAGE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CLASS ATTENDED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
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
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ATTENDENCE"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3600
      TabIndex        =   0
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command5_Click()
Text4.Text = Val(Text3.Text) / 60 * 100
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
Form3.Hide
Form2.Show
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveNext

End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MovePrevious

End Sub

Private Sub Command8_Click()
DataReport1.Show

End Sub
