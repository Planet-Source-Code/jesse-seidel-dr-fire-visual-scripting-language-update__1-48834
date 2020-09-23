VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Components"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1290
   Icon            =   "toolbar.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   1290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VectorBasic.chameleonButton cb1 
      Height          =   435
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   ""
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
      MICON           =   "toolbar.frx":030A
      PICN            =   "toolbar.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   600
      TabIndex        =   24
      Text            =   "True"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.TextBox PicText1 
      Height          =   285
      Left            =   600
      TabIndex        =   22
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Set Values"
      Height          =   315
      Left            =   120
      TabIndex        =   20
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   720
      TabIndex        =   19
      Text            =   "New Application"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   600
      TabIndex        =   17
      Text            =   "4000"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   600
      TabIndex        =   15
      Text            =   "6000"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   12
      Text            =   "True"
      Top             =   4080
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Text            =   "New Caption"
      Top             =   4320
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   600
      TabIndex        =   10
      Text            =   "200"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   600
      TabIndex        =   9
      Text            =   "1000"
      Top             =   3600
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   360
      TabIndex        =   8
      Text            =   "0"
      Top             =   3360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Text            =   "0"
      Top             =   3120
      Width           =   1335
   End
   Begin VectorBasic.chameleonButton cb2 
      Height          =   435
      Left            =   720
      TabIndex        =   26
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   ""
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
      MICON           =   "toolbar.frx":0778
      PICN            =   "toolbar.frx":0794
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VectorBasic.chameleonButton chameleonButton1 
      Height          =   435
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   ""
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
      MICON           =   "toolbar.frx":0C26
      PICN            =   "toolbar.frx":0C42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VectorBasic.chameleonButton chameleonButton2 
      Height          =   435
      Left            =   720
      TabIndex        =   28
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      BTYPE           =   8
      TX              =   ""
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
      MICON           =   "toolbar.frx":11D4
      PICN            =   "toolbar.frx":11F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label18 
      Caption         =   "Stretch:"
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label17 
      Caption         =   "Picture:"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   4560
      Width           =   615
   End
   Begin VB.Label Label16 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   0
      TabIndex        =   18
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Height:"
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "Width:"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5400
      Width           =   735
   End
   Begin VB.Label Label13 
      Caption         =   "Form Properties:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label12 
      Caption         =   "Caption:"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "Visible:"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "Height:"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "Width:"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label8 
      Caption         =   "Top:"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label7 
      Caption         =   "Left:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Properties:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo SpawnCmdError:
If Combo1.Text = "True" Then
CreateCommandButton Form1.Command1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
ElseIf Combo1.Text = "False" Then
CreateCommandButton Form1.Command1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnCmdError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub cb1_Click()
On Error GoTo SpawnCmdError:
If Combo1.Text = "True" Then
CreateCommandButton Form1.Command1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
CreateFrame Form6.Frame2, 0, 0, Form6.Frame2.Width, Form6.Frame2.Height, True, Form1.Command1(Index).Name
ElseIf Combo1.Text = "False" Then
CreateCommandButton Form1.Command1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnCmdError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub cb2_Click()
On Error GoTo SpawnTextError:
If Combo1.Text = "True" Then
CreateTextBox Form1.Text1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
ElseIf Combo1.Text = "False" Then
CreateTextBox Form1.Text1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnTextError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub Command2_Click()
On Error GoTo SpawnTextError:
If Combo1.Text = "True" Then
CreateTextBox Form1.Text1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
ElseIf Combo1.Text = "False" Then
CreateTextBox Form1.Text1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnTextError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub chameleonButton1_Click()
On Error GoTo SpawnLabelError:
If Combo1.Text = "True" Then
CreateLabel Form1.Label1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
ElseIf Combo1.Text = "False" Then
CreateLabel Form1.Label1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnLabelError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub Command3_Click()
On Error GoTo SpawnLabelError:
If Combo1.Text = "True" Then
CreateLabel Form1.Label1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, Text5.Text
ElseIf Combo1.Text = "False" Then
CreateLabel Form1.Label1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, Text5.Text
End If
Exit Sub
SpawnLabelError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub chameleonButton2_Click()
On Error GoTo SpawnImageError:
NewPicPath Form1.PicPath, PicText1.Text
If Combo1.Text = "True" Then
    If Combo2.Text = "True" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, PicText1.Text, True
    ElseIf Combo2.Text = "False" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, PicText1.Text, False
    End If
ElseIf Combo1.Text = "False" Then
    If Combo2.Text = "True" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, PicText1.Text, True
    ElseIf Combo2.Text = "False" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, PicText1.Text, False
    End If
End If
Exit Sub
SpawnImageError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub Command4_Click()
On Error GoTo SpawnImageError:
NewPicPath Form1.PicPath, PicText1.Text
If Combo1.Text = "True" Then
    If Combo2.Text = "True" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, PicText1.Text, True
    ElseIf Combo2.Text = "False" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, True, PicText1.Text, False
    End If
ElseIf Combo1.Text = "False" Then
    If Combo2.Text = "True" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, PicText1.Text, True
    ElseIf Combo2.Text = "False" Then
CreateImageBox Form1.Image1, Text1.Text, Text2.Text, Text3.Text, Text4.Text, False, PicText1.Text, False
    End If
End If
Exit Sub
SpawnImageError:
MsgBox "Error, please check settings and retry!", , "Error"
End Sub

Private Sub Command5_Click()
On Error GoTo FormSetErr:
Form1.Width = Text6.Text
Form1.Height = Text7.Text
Form1.Caption = Text8.Text
Exit Sub
FormSetErr:
MsgBox "Error, the form cannot take these settings. Please change them.", , "Error"
End Sub

Private Sub Form_Load()
On Error GoTo ToolbarLoadError:
Form1.Show
Form3.Show
Form3.Left = Form2.Left + 1815
Form3.Top = Form2.Top
Form1.Top = Form2.Top + 6945 - 2670
Form1.Left = Form2.Left + 1815
Form5.Show
Form5.Left = Form2.Left + Form2.Width + Form3.Width
Form5.Top = Me.Top + 3510
V.Show
V.Left = Form3.Left + 5250
V.Top = Form3.Top

Combo1.AddItem "True", 0
Combo1.AddItem "False", 1

Combo2.AddItem "True", 0
Combo2.AddItem "False", 1
Exit Sub
ToolbarLoadError:
MsgBox "Toolbar loading error occurred! Contact me about this please.", , "Error"
End Sub

