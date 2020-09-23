VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1Caption"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add!"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check?"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   0
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Age:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
