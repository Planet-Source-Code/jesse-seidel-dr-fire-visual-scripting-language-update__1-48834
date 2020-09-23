VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form Translator"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4080
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   2775
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   2280
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label Label2 
      Caption         =   "InterBasic Environment Code:"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Visual Basic Form Code:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Visual Basic Form"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save InterBasic Form"
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
If Form1.Height > 600 Then
Form1.Height = 3900
Form1.Width = 4710
End If
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuFileOpen_Click()
On Error Resume Next
Dim FormWidthF As Integer
Dim FormHeightF As Integer
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Visual Basic Form File|*.frm"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Input As 1
If Err.Number = 75 Then
Err.Number = 0
Exit Sub
End If
Text1.Text = ""
Text2.Text = ""
Do Until EOF(1)
Line Input #1, CurLineFromFile
Text1.Text = Text1.Text & CurLineFromFile & vbCrLf
Loop
Close 1

EnvironmentCode = Text1.Text
EnvironmentCode1 = Text1.Text

T = 0
C = 0
L = 0

FormCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.Form ") + 6)
FormCode = Mid(FormCode, 1, InStr(1, FormCode, "Begin") - 1)
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.Form ") + 6)

FormCaption = Mid(FormCode, InStr(1, FormCode, "Caption"))
FormCaption = Mid(FormCaption, 1, InStr(1, FormCaption, vbCrLf))
FormCaption = Replace(FormCaption, "Caption", "")
FormCaption = Replace(FormCaption, Chr(13), "")
FormCaption = Replace(FormCaption, " ", "")
FormCaption = Replace(FormCaption, Chr(34), "")
FormCaption = Replace(FormCaption, "=", "")

FormHeight = Mid(FormCode, InStr(1, FormCode, "ClientHeight"))
FormHeight = Mid(FormHeight, 1, InStr(1, FormHeight, vbCrLf))
FormHeight = Replace(FormHeight, "ClientHeight", "")
FormHeight = Replace(FormHeight, Chr(13), "")
FormHeight = Replace(FormHeight, " ", "")
FormHeight = Replace(FormHeight, Chr(34), "")
FormHeight = Replace(FormHeight, "=", "")
FormHeightF = FormHeight
FormHeightF = FormHeightF + 500

FormWidth = Mid(FormCode, InStr(1, FormCode, "ClientWidth"))
FormWidth = Mid(FormWidth, 1, InStr(1, FormWidth, vbCrLf))
FormWidth = Replace(FormWidth, "ClientWidth", "")
FormWidth = Replace(FormWidth, Chr(13), "")
FormWidth = Replace(FormWidth, " ", "")
FormWidth = Replace(FormWidth, Chr(34), "")
FormWidth = Replace(FormWidth, "=", "")
FormWidthF = FormWidth
FormWidthF = FormWidthF + 120

Text2.Text = ""
Text2.Text = Text2.Text & "Form()  " & "Width:" & FormWidthF & " Height:" & FormHeightF & " Caption:" & FormCaption & vbCrLf

T = 0

Do While InStr(1, EnvironmentCode, "Begin VB.TextBox") <> 0

T = T + 1
TextBCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.TextBox ") + 17)
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.TextBox ") + 6)

TextBCode = Mid(TextBCode, 1, InStr(1, TextBCode, "Begin") - 1)
TextBCaption = Mid(TextBCode, InStr(1, TextBCode, "Text            =   " & Chr(34)))
TextBCaption = Mid(TextBCaption, 1, InStr(1, TextBCaption, vbCrLf))
TextBCaption = Replace(TextBCaption, Chr(13), "")
TextBCaption = Replace(TextBCaption, " ", "")
TextBCaption = Replace(TextBCaption, Chr(34), "")
TextBCaption = Replace(TextBCaption, "=", "")
TextBCaption = Replace(TextBCaption, "TextBox", "")
TextBCaption = Mid(TextBCaption, 5)

TextBHeight = Mid(TextBCode, InStr(1, TextBCode, "Height"))
TextBHeight = Mid(TextBHeight, 1, InStr(1, TextBHeight, vbCrLf))
TextBHeight = Replace(TextBHeight, "Height", "")
TextBHeight = Replace(TextBHeight, Chr(13), "")
TextBHeight = Replace(TextBHeight, " ", "")
TextBHeight = Replace(TextBHeight, Chr(34), "")
TextBHeight = Replace(TextBHeight, "=", "")

TextBWidth = Mid(TextBCode, InStr(1, TextBCode, "Width"))
TextBWidth = Mid(TextBWidth, 1, InStr(1, TextBWidth, vbCrLf))
TextBWidth = Replace(TextBWidth, "Width", "")
TextBWidth = Replace(TextBWidth, Chr(13), "")
TextBWidth = Replace(TextBWidth, " ", "")
TextBWidth = Replace(TextBWidth, Chr(34), "")
TextBWidth = Replace(TextBWidth, "=", "")

TextBLeft = Mid(TextBCode, InStr(1, TextBCode, "Left"))
TextBLeft = Mid(TextBLeft, 1, InStr(1, TextBLeft, vbCrLf))
TextBLeft = Replace(TextBLeft, "Left", "")
TextBLeft = Replace(TextBLeft, Chr(13), "")
TextBLeft = Replace(TextBLeft, " ", "")
TextBLeft = Replace(TextBLeft, Chr(34), "")
TextBLeft = Replace(TextBLeft, "=", "")

TextBTop = Mid(TextBCode, InStr(1, TextBCode, "Top"))
TextBTop = Mid(TextBTop, 1, InStr(1, TextBTop, vbCrLf))
TextBTop = Replace(TextBTop, "Top", "")
TextBTop = Replace(TextBTop, Chr(13), "")
TextBTop = Replace(TextBTop, " ", "")
TextBTop = Replace(TextBTop, Chr(34), "")
TextBTop = Replace(TextBTop, "=", "")
On Error Resume Next
TextBFalse = Mid(TextBCode, InStr(1, TextBCode, "Visible"))
TextBFalse = Mid(TextBTop, 1, InStr(1, TextBFalse, vbCrLf))
TextBFalse = Replace(TextBFalse, "Visible", "")
TextBFalse = Replace(TextBFalse, Chr(13), "")
TextBFalse = Replace(TextBFalse, " ", "")
TextBFalse = Replace(TextBFalse, Chr(34), "")
TextBFalse = Replace(TextBFalse, "=", "")
TextBFalse = Replace(TextBFalse, "'", "")
TextBFalse = Replace(TextBFalse, "0", "")
If Err.Number = 5 Then
Err.Number = 0
TextBFalse = "True"
End If

Text2.Text = Text2.Text & "CreateCtrl Text(" & T & ") Left:" & TextBLeft & " Top:" & TextBTop & " Width:" & TextBWidth & " Height:" & TextBHeight & " Visible:" & TextBFalse & " Caption:" & TextBCaption & vbCrLf
Loop

EnvironmentCode = EnvironmentCode1

C = 0
Do While InStr(1, EnvironmentCode, "Begin VB.CommandButton") <> 0
C = C + 1
CmdBCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.CommandButton ") + 23)
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.CommandButton ") + 6)

CmdBCode = Mid(CmdBCode, 1, InStr(1, CmdBCode, "Begin") - 1)
CmdBCaption = Mid(CmdBCode, InStr(1, CmdBCode, "Caption         =   " & Chr(34)))
CmdBCaption = Mid(CmdBCaption, 1, InStr(1, CmdBCaption, vbCrLf))
CmdBCaption = Replace(CmdBCaption, Chr(13), "")
CmdBCaption = Replace(CmdBCaption, " ", "")
CmdBCaption = Replace(CmdBCaption, Chr(34), "")
CmdBCaption = Replace(CmdBCaption, "=", "")
CmdBCaption = Mid(CmdBCaption, 8)

CmdBHeight = Mid(CmdBCode, InStr(1, CmdBCode, "Height"))
CmdBHeight = Mid(CmdBHeight, 1, InStr(1, CmdBHeight, vbCrLf))
CmdBHeight = Replace(CmdBHeight, "Height", "")
CmdBHeight = Replace(CmdBHeight, Chr(13), "")
CmdBHeight = Replace(CmdBHeight, " ", "")
CmdBHeight = Replace(CmdBHeight, Chr(34), "")
CmdBHeight = Replace(CmdBHeight, "=", "")

CmdBWidth = Mid(CmdBCode, InStr(1, CmdBCode, "Width"))
CmdBWidth = Mid(CmdBWidth, 1, InStr(1, CmdBWidth, vbCrLf))
CmdBWidth = Replace(CmdBWidth, "Width", "")
CmdBWidth = Replace(CmdBWidth, Chr(13), "")
CmdBWidth = Replace(CmdBWidth, " ", "")
CmdBWidth = Replace(CmdBWidth, Chr(34), "")
CmdBWidth = Replace(CmdBWidth, "=", "")

CmdBLeft = Mid(CmdBCode, InStr(1, CmdBCode, "Left"))
CmdBLeft = Mid(CmdBLeft, 1, InStr(1, CmdBLeft, vbCrLf))
CmdBLeft = Replace(CmdBLeft, "Left", "")
CmdBLeft = Replace(CmdBLeft, Chr(13), "")
CmdBLeft = Replace(CmdBLeft, " ", "")
CmdBLeft = Replace(CmdBLeft, Chr(34), "")
CmdBLeft = Replace(CmdBLeft, "=", "")

CmdBTop = Mid(CmdBCode, InStr(1, CmdBCode, "Top"))
CmdBTop = Mid(CmdBTop, 1, InStr(1, CmdBTop, vbCrLf))
CmdBTop = Replace(CmdBTop, "Top", "")
CmdBTop = Replace(CmdBTop, Chr(13), "")
CmdBTop = Replace(CmdBTop, " ", "")
CmdBTop = Replace(CmdBTop, Chr(34), "")
CmdBTop = Replace(CmdBTop, "=", "")
On Error Resume Next
CmdBFalse = Mid(CmdBCode, InStr(1, CmdBCode, "Visible"))
CmdBFalse = Mid(CmdBTop, 1, InStr(1, CmdBFalse, vbCrLf))
CmdBFalse = Replace(CmdBFalse, "Visible", "")
CmdBFalse = Replace(CmdBFalse, Chr(13), "")
CmdBFalse = Replace(CmdBFalse, " ", "")
CmdBFalse = Replace(CmdBFalse, Chr(34), "")
CmdBFalse = Replace(CmdBFalse, "=", "")
CmdBFalse = Replace(CmdBFalse, "'", "")
CmdBFalse = Replace(CmdBFalse, "0", "")
If Err.Number = 5 Then
Err.Number = 0
CmdBFalse = "True"
End If

Text2.Text = Text2.Text & "CreateCtrl Cmd(" & C & ") Left:" & CmdBLeft & " Top:" & CmdBTop & " Width:" & CmdBWidth & " Height:" & CmdBHeight & " Visible:" & CmdBFalse & " Caption:" & CmdBCaption & vbCrLf
Loop

EnvironmentCode = EnvironmentCode1

L = 0
Do While InStr(1, EnvironmentCode, "Begin VB.Label") <> 0
L = L + 1
LabelBCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.Label ") + 15)
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, "Begin VB.Label ") + 6)

LabelBCode = Mid(LabelBCode, 1, InStr(1, LabelBCode, "Begin") - 1)
LabelBCaption = Mid(LabelBCode, InStr(1, LabelBCode, "Caption         =   " & Chr(34)))
LabelBCaption = Mid(LabelBCaption, 1, InStr(1, LabelBCaption, vbCrLf))
LabelBCaption = Replace(LabelBCaption, Chr(13), "")
LabelBCaption = Replace(LabelBCaption, " ", "")
LabelBCaption = Replace(LabelBCaption, Chr(34), "")
LabelBCaption = Replace(LabelBCaption, "=", "")
LabelBCaption = Mid(LabelBCaption, 8)

LabelBHeight = Mid(LabelBCode, InStr(1, LabelBCode, "Height"))
LabelBHeight = Mid(LabelBHeight, 1, InStr(1, LabelBHeight, vbCrLf))
LabelBHeight = Replace(LabelBHeight, "Height", "")
LabelBHeight = Replace(LabelBHeight, Chr(13), "")
LabelBHeight = Replace(LabelBHeight, " ", "")
LabelBHeight = Replace(LabelBHeight, Chr(34), "")
LabelBHeight = Replace(LabelBHeight, "=", "")

LabelBWidth = Mid(LabelBCode, InStr(1, LabelBCode, "Width"))
LabelBWidth = Mid(LabelBWidth, 1, InStr(1, LabelBWidth, vbCrLf))
LabelBWidth = Replace(LabelBWidth, "Width", "")
LabelBWidth = Replace(LabelBWidth, Chr(13), "")
LabelBWidth = Replace(LabelBWidth, " ", "")
LabelBWidth = Replace(LabelBWidth, Chr(34), "")
LabelBWidth = Replace(LabelBWidth, "=", "")

LabelBLeft = Mid(LabelBCode, InStr(1, LabelBCode, "Left"))
LabelBLeft = Mid(LabelBLeft, 1, InStr(1, LabelBLeft, vbCrLf))
LabelBLeft = Replace(LabelBLeft, "Left", "")
LabelBLeft = Replace(LabelBLeft, Chr(13), "")
LabelBLeft = Replace(LabelBLeft, " ", "")
LabelBLeft = Replace(LabelBLeft, Chr(34), "")
LabelBLeft = Replace(LabelBLeft, "=", "")

LabelBTop = Mid(LabelBCode, InStr(1, LabelBCode, "Top"))
LabelBTop = Mid(LabelBTop, 1, InStr(1, LabelBTop, vbCrLf))
LabelBTop = Replace(LabelBTop, "Top", "")
LabelBTop = Replace(LabelBTop, Chr(13), "")
LabelBTop = Replace(LabelBTop, " ", "")
LabelBTop = Replace(LabelBTop, Chr(34), "")
LabelBTop = Replace(LabelBTop, "=", "")
On Error Resume Next
LabelBFalse = Mid(LabelBCode, InStr(1, LabelBCode, "Visible"))
LabelBFalse = Mid(LabelBTop, 1, InStr(1, LabelBFalse, vbCrLf))
LabelBFalse = Replace(LabelBFalse, "Visible", "")
LabelBFalse = Replace(LabelBFalse, Chr(13), "")
LabelBFalse = Replace(LabelBFalse, " ", "")
LabelBFalse = Replace(LabelBFalse, Chr(34), "")
LabelBFalse = Replace(LabelBFalse, "=", "")
LabelBFalse = Replace(LabelBFalse, "'", "")
LabelBFalse = Replace(LabelBFalse, "0", "")
If Err.Number = 5 Then
Err.Number = 0
LabelBFalse = "True"
End If

Text2.Text = Text2.Text & "CreateCtrl Label(" & L & ") Left:" & LabelBLeft & " Top:" & LabelBTop & " Width:" & LabelBWidth & " Height:" & LabelBHeight & " Visible:" & LabelBFalse & " Caption:" & LabelBCaption & vbCrLf
Loop

End Sub

Private Sub mnuFileSave_Click()
On Error Resume Next
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Environment Code File|*.ibf"
CommonDialog1.ShowSave
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Output As 1
If Err.Number = 75 Then
Err.Number = 0
Exit Sub
End If
Print #1, Text2.Text
Close 1
End Sub





















