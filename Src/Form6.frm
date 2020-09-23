VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8010
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Button"
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "Form6.frx":0000
         Left            =   600
         List            =   "Form6.frx":002B
         TabIndex        =   27
         Text            =   "2 - Windows 32-bit"
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   26
         Text            =   "200"
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   19
         Text            =   "0"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Text            =   "0"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   17
         Text            =   "Text"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Text            =   "1000"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form6.frx":0101
         Left            =   720
         List            =   "Form6.frx":010B
         TabIndex        =   15
         Text            =   "True"
         Top             =   1440
         Width           =   1695
      End
      Begin VectorBasic.chameleonButton chameleonButton2 
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2160
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Update Properties"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160664
         BCOLO           =   13160664
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form6.frx":011C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         Caption         =   "Style:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label11 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Visible:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6240
      TabIndex        =   7
      Text            =   "200"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "TextBox"
      Height          =   2175
      Left            =   5520
      TabIndex        =   0
      Top             =   1440
      Width           =   2175
      Begin VectorBasic.chameleonButton chameleonButton1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "Update Properties"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160664
         BCOLO           =   13160664
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Form6.frx":0138
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "Form6.frx":0154
         Left            =   720
         List            =   "Form6.frx":015E
         TabIndex        =   12
         Text            =   "True"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   8
         Text            =   "1000"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   6
         Text            =   "Text"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   2
         Text            =   "0"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Visible:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Height:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Left:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Top:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo3_Change()
Text11.Text = Combo3.Text
Text11.Text = Replace(Text11.Text, "1 - ", "")
Text11.Text = Replace(Text11.Text, "2 - ", "")
Text11.Text = Replace(Text11.Text, "3 - ", "")
Text11.Text = Replace(Text11.Text, "4 - ", "")
Text11.Text = Replace(Text11.Text, "5 - ", "")
Text11.Text = Replace(Text11.Text, "6 - ", "")
Text11.Text = Replace(Text11.Text, "7 - ", "")
Text11.Text = Replace(Text11.Text, "8 - ", "")
Text11.Text = Replace(Text11.Text, "9 - ", "")
Text11.Text = Replace(Text11.Text, "10 - ", "")
Text11.Text = Replace(Text11.Text, "11 - ", "")
Text11.Text = Replace(Text11.Text, "12 - ", "")
Text11.Text = Replace(Text11.Text, "13 - ", "")
exbtn.ButtonType = Text11.Text
Label15.Caption = Text11.Text
End Sub

Private Sub Command1_Click()
Form1.Command1(2).Caption = "y-"
End Sub
