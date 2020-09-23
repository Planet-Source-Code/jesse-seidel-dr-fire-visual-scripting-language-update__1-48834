VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "New Application"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   Icon            =   "runtime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox PicPath 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Index           =   0
      Left            =   3600
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   495
      Index           =   0
      Left            =   2400
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CodeExecute As String
Dim EnvironmentCode As String

Dim FormAnalysis As String
Dim formwidthf As Integer
Dim FormHeightF As Integer
Dim FormCaption As String

Dim CurName As String
Dim CurLeftF As Integer
Dim CurTopF As Integer
Dim CurWidthF As Integer
Dim CurHeightF As Integer
Dim CurVisible As String
Dim CurCaption As String
Dim CurNumberF As Integer
Dim CurPicture As String

Private Sub Command1_Click(Index As Integer)
On Error GoTo CmdExecError:
'Check Code Exists
If InStr(1, Form3.Text1.Text, "Start Cmd(" & Index & ")_Clicked()") = 0 Then
Exit Sub
End If

'Get Code
GetCode = Mid(Form3.Text1.Text, InStr(1, Form3.Text1.Text, "Start Cmd(" & Index & ")_Clicked()") + Len("Start Cmd(" & Index & ")_Clicked()") + 2)
GetCode = Mid(GetCode, 1, InStr(1, GetCode, "End Cmd(" & Index & ")_Clicked") - 3)

'Execute Code
ExecuteAppCode (GetCode)
Exit Sub
CmdExecError:
MsgBox "The code couldn't be executed, please check that there is any!", , "Error"
End Sub

Private Sub Form_Load()
On Error GoTo RuntimeError:
EnvironmentCode = Form3.Text2.Text
CodeExecute = Form3.Text1.Text

'Construct Form
If InStr(1, EnvironmentCode, "Form()") = 0 Then
Unload Me
Exit Sub
ElseIf InStr(1, EnvironmentCode, "Form()") > 3 Then
MsgBox "Form Environment Code must come first, before any control Environment Code.", , "Error"
Unload Me
Exit Sub
ElseIf InStr(1, EnvironmentCode, "Form()") < 3 Then

'Get Form Information
FormAnalysis = Mid(EnvironmentCode, 1, InStr(1, EnvironmentCode, vbCrLf))
Formwidth = Mid(FormAnalysis, InStr(1, FormAnalysis, "Width:") + 6)
Formwidth = Mid(Formwidth, 1, InStr(1, Formwidth, " ") - 1)
formwidthf = Formwidth
FormHeight = Mid(FormAnalysis, InStr(1, FormAnalysis, "Height:") + 7)
FormHeight = Mid(FormHeight, 1, InStr(1, FormHeight, " ") - 1)
FormHeightF = FormHeight
FormCaption = Mid(FormAnalysis, InStr(1, FormAnalysis, "Caption:") + 8)
FormCaption = Mid(FormCaption, 1, InStr(1, FormCaption, Chr(13)) - 1)

'Construct Form
Form4.Width = formwidthf
Form4.Height = FormHeightF
Form4.Caption = FormCaption
End If

Do While Len(EnvironmentCode) > 2
CurControlSpecs = Mid(EnvironmentCode, 1, InStr(1, EnvironmentCode, vbCrLf))
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, vbCrLf) + 2)
CurControlSpecs = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Create Ctrl ") + 12)

If InStr(1, CurControlSpecs, "Image") >= 1 Then
'Number
CurNumber = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "(") + 1)
CurNumber = Mid(CurNumber, 1, InStr(1, CurNumber, "(") + 1)
CurNumberF = CurNumber
NewPicPath Form4.PicPath, ""
End If

'Name
If InStr(1, CurControlSpecs, "Cmd") = 1 Then
CurName = "Cmd"
ElseIf InStr(1, CurControlSpecs, "Text") = 1 Then
CurName = "Text"
ElseIf InStr(1, CurControlSpecs, "Label") = 1 Then
CurName = "Label"
ElseIf InStr(1, CurControlSpecs, "Image") = 1 Then
CurName = "Image"
End If


'Left
CurLeft = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Left:") + 5)
CurLeft = Mid(CurLeft, 1, InStr(1, CurLeft, " ") - 1)
CurLeftF = CurLeft

'Top
CurTop = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Top:") + 4)
CurTop = Mid(CurTop, 1, InStr(1, CurTop, " ") - 1)
CurTopF = CurTop

'Width
CurWidthF = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Width:") + 6)
CurWidthF = Mid(CurWidthF, 1, InStr(1, CurWidthF, " ") - 1)
CurWidthF = CurWidthF

'Height
CurHeightF = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Height:") + 7)
CurHeightF = Mid(CurHeightF, 1, InStr(1, CurHeightF, " ") - 1)
CurHeightF = CurHeightF

'Visible
CurVisible = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Visible:") + 8)
CurVisible = Mid(CurVisible, 1, InStr(1, CurVisible, " ") - 1)

'Caption
If InStr(1, CurControlSpecs, "Caption:") <> 0 Then
CurCaption = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Caption:") + 8)
CurCaption = Mid(CurCaption, 1, InStr(1, CurCaption, Chr(13)) - 1)
End If

If CurName = "Image" Then
'Picture
CurPicture = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Picture:") + 8)
CurPicture = Mid(CurPicture, 1, InStr(1, CurPicture, " Stretch:") - 1)
PicPath(CurNumber).Text = CurPicture

'Stretch
CurStretch = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Stretch") + 8)
CurStretch = Mid(CurStretch, 1, InStr(1, CurStretch, Chr(13)) - 1)
End If

If CurName = "Cmd" Then
    If CurVisible = "True" Then
CreateCommandButton Form4.Command1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateCommandButton Form4.Command1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Text" Then
    If CurVisible = "True" Then
CreateTextBox Form4.Text1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateTextBox Form4.Text1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Label" Then
    If CurVisible = "True" Then
CreateLabel Form4.Label1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateLabel Form4.Label1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Image" Then
    If CurVisible = "True" Then
        If CurStretch = "True" Then
CreateImageBox Form4.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, PicPath(CurNumberF).Text, True
        ElseIf CurStretch = "False" Then
CreateImageBox Form4.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, PicPath(CurNumberF).Text, False
        End If
    ElseIf CurVisible = "False" Then
        If CurStretch = "True" Then
CreateImageBox Form4.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, PicPath(CurNumberF).Text, True
        ElseIf CurStretch = "False" Then
CreateImageBox Form4.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, PicPath(CurNumberF).Text, False
        End If
    End If
End If

Loop
Exit Sub
RuntimeError:
MsgBox "The form couldn't be loaded or the code was incorrect. Please check everything, then contact me if the error persists!", , "Error"
End Sub
