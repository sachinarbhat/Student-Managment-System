VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H80000007&
   Caption         =   "Form9"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13800
   LinkTopic       =   "Form9"
   Picture         =   "querry2.frx":0000
   ScaleHeight     =   6105
   ScaleWidth      =   13800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "BACK"
      Height          =   495
      Left            =   10440
      TabIndex        =   2
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MARKS1<50"
      Height          =   735
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PERCENTAGE greater than 70"
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   1320
      Width           =   3015
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.RecordSource = "SELECT USN,TOTAL,PERCENTAGE FROM RESULT WHERE PERCENTAGE>70"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource
End Sub

Private Sub Command2_Click()
Adodc1.RecordSource = "SELECT USN,MARKS1 FROM RESULT WHERE MARKS1<50"
Adodc1.Refresh
Adodc1.Caption = Adodc1.RecordSource

End Sub

Private Sub Command3_Click()
Form9.Hide
Form5.Show
End Sub
