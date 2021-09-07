VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13575
   LinkTopic       =   "Form8"
   Picture         =   "querry1.frx":0000
   ScaleHeight     =   6060
   ScaleWidth      =   13575
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   495
      Left            =   10200
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "HAVING ME"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SELECT FEMALE"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "SELECT USN,NAME,GENDER FROM STUDENT WHERE GENDER = 'F' "
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "SELECT NAME,DOB FROM STUDENT WHERE DID = 'ME'"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource

End Sub

Private Sub Command3_Click()
Form8.Hide
Form4.Show
End Sub
