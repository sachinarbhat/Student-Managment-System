VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16275
   LinkTopic       =   "Form5"
   Picture         =   "Result.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command9 
      Caption         =   "Querry"
      Height          =   615
      Left            =   14520
      TabIndex        =   21
      Top             =   6600
      Width           =   1695
   End
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
      Height          =   615
      Left            =   12120
      TabIndex        =   20
      Top             =   6600
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "PREVIOUS"
      Height          =   375
      Left            =   14040
      TabIndex        =   19
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "NEXT"
      Height          =   375
      Left            =   11640
      TabIndex        =   18
      Top             =   5880
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9960
      TabIndex        =   17
      Top             =   4560
      Width           =   975
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
      Left            =   9600
      TabIndex        =   16
      Top             =   6600
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
      Height          =   615
      Left            =   6720
      TabIndex        =   15
      Top             =   6600
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
      Left            =   3720
      TabIndex        =   14
      Top             =   6600
      Width           =   2295
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
      Left            =   840
      TabIndex        =   13
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox Text6 
      DataField       =   "Percentage"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6600
      TabIndex        =   12
      Top             =   4920
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      DataField       =   "Total"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   0
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6600
      TabIndex        =   11
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text4 
      DataField       =   "Marks3"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6600
      TabIndex        =   10
      Top             =   3600
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      DataField       =   "Marks2"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   6600
      TabIndex        =   9
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      DataField       =   "Marks1"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   6600
      TabIndex        =   8
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      DataField       =   "USN"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   6600
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label7 
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
      Left            =   3120
      TabIndex        =   6
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL MARKS"
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
      Left            =   3120
      TabIndex        =   5
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MARKS 3"
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
      Left            =   3120
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MARKS 2"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MARKS 1"
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
      Left            =   3120
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "RESULT"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   10215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command5_Click()

Text5.Text = Val(Text2.Text) + Val(Text3.Text) + Val(Text4.Text)
Text6.Text = (Val(Text5.Text) / 300) * 100



End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MovePrevious
End Sub

Private Sub Command8_Click()
Form5.Hide

DataReport2.Show

End Sub

Private Sub Command9_Click()
Form5.Hide
Form9.Show
End Sub
