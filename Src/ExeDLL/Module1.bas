Attribute VB_Name = "Module1"
Public Sub Pause(Duration As Long)
    Dim Current As Long
    Current = Timer
    Do Until Timer - Current >= Duration
        DoEvents
    Loop
End Sub

Public Function CreateCommandButton(Control As Variant, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Visible As Boolean, Caption As String)
On Error GoTo CmdError:
NewControl = Control.Count
Load Control(NewControl)
With Control(NewControl)
    .Left = Left
    .Top = Top
    .Width = Width
    .Height = Height
    .Visible = Visible
    .Caption = Caption
End With
Exit Function
CmdError:
MsgBox "An error creating the command button occurred! Please try again, if you cannot fix this, please contact me.", , "Error"
End Function

Public Function CreateLabel(Control As Variant, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Visible As Boolean, Caption As String)
On Error GoTo LabelError:
NewControl = Control.Count
Load Control(NewControl)
With Control(NewControl)
    .Left = Left
    .Top = Top
    .Width = Width
    .Height = Height
    .Visible = Visible
    .Caption = Caption
End With
Exit Function
LabelError:
MsgBox "An error creating the label control occurred! Please try again, if you cannot fix this, please contact me.", , "Error"
End Function

Public Function CreateTextBox(Control As Variant, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Visible As Boolean, Caption As String)
On Error GoTo TextBoxError:
NewControl = Control.Count
Load Control(NewControl)
With Control(NewControl)
    .Left = Left
    .Top = Top
    .Width = Width
    .Height = Height
    .Visible = Visible
    .Text = Caption
End With
Exit Function
TextBoxError:
MsgBox "An error creating the text box occurred! Please try again, if you cannot fix this, please contact me.", , "Error"
End Function

Public Function CreateImageBox(Control As Variant, Left As Integer, Top As Integer, Width As Integer, Height As Integer, Visible As Boolean, Picture As String, Stretch As Boolean)
On Error GoTo ImageError:
NewControl = Control.Count
Load Control(NewControl)
With Control(NewControl)
    .Left = Left
    .Top = Top
    .Width = Width
    .Height = Height
    .Visible = Visible
    .Stretch = Stretch
If Len(Picture) <> 0 Then
    .Picture = LoadPicture(Picture)
End If
End With
Exit Function
ImageError:
MsgBox "An error creating the image box control occurred! Please try again, if you cannot fix this, please contact me.", , "Error"
End Function

Public Function NewPicPath(Control As Variant, Path As String)
On Error GoTo PicPathError:
NewControl = Control.Count
Load Control(NewControl)
With Control(NewControl)
    .Text = Path
End With
Exit Function
PicPathError:
MsgBox "An error creating the image box path variable occurred! Please try again, if you cannot fix this, please contact me.", , "Error"
End Function


Public Function BuildEnvironmentCode(Control As Variant, Name As String, Form As Form)
On Error GoTo BuildEnvError:
Dim X As Integer
TheCount = Control.Count
For X = 2 To TheCount
If Name = "Command1" Then
Form3.Text2.Text = Form3.Text2.Text & "CreateCtrl Cmd(" & X - 1 & ")" & " Left:" & Control(X - 1).Left & " Top:" & Control(X - 1).Top & " Width:" & Control(X - 1).Width & " Height:" & Control(X - 1).Height & " Visible:" & Control(X - 1).Visible & " Caption:" & Control(X - 1).Caption & vbCrLf
ElseIf Name = "Text1" Then
Form3.Text2.Text = Form3.Text2.Text & "CreateCtrl Text(" & X - 1 & ")" & " Left:" & Control(X - 1).Left & " Top:" & Control(X - 1).Top & " Width:" & Control(X - 1).Width & " Height:" & Control(X - 1).Height & " Visible:" & Control(X - 1).Visible & " Caption:" & Control(X - 1).Text & vbCrLf
ElseIf Name = "Label1" Then
Form3.Text2.Text = Form3.Text2.Text & "CreateCtrl Label(" & X - 1 & ")" & " Left:" & Control(X - 1).Left & " Top:" & Control(X - 1).Top & " Width:" & Control(X - 1).Width & " Height:" & Control(X - 1).Height & " Visible:" & Control(X - 1).Visible & " Caption:" & Control(X - 1).Caption & vbCrLf
ElseIf Name = "Image1" Then
PicPath = Form.PicPath(X - 1).Text
Form3.Text2.Text = Form3.Text2.Text & "CreateCtrl Image(" & X - 1 & ")" & " Left:" & Control(X - 1).Left & " Top:" & Control(X - 1).Top & " Width:" & Control(X - 1).Width & " Height:" & Control(X - 1).Height & " Visible:" & Control(X - 1).Visible & " Picture:" & PicPath & " Stretch:" & Control(X - 1).Stretch & vbCrLf
End If
Next X
Exit Function
BuildEnvError:
MsgBox "An error building the environment code occurred. Please try to fix this, if you cannot, please contact me.", , "Error"
End Function

Public Function BuildEnvironmentFormCode(Form As Form)
On Error GoTo BuildFrmError:
Form3.Text2.Text = Form3.Text2.Text & "Form() " & " Width:" & Form.Width & " Height:" & Form.Height & " Caption:" & Form.Caption & vbCrLf
Exit Function
BuildFrmError:
MsgBox "An error building the form environment code occurred. Please try to fix this, if you cannot, please contact me.", , "Error"
End Function

Public Function LoadFromFile(TextBox As TextBox, Path As String)
On Error GoTo LoadFileError:
Open Path For Input As 1
Do Until EOF(1)
Line Input #1, CurLineFromFile
TextBox.Text = TextBox.Text & CurLineFromFile & vbCrLf
Loop
Close 1
Exit Function
LoadFileError:
MsgBox "An error occurred with the textbox's 'LoadFile' command. Does the file exist?", , "Error"
End Function

Public Function SaveToFile(TextBox As TextBox, Path As String)
On Error GoTo SaveFileError:
Open Path For Output As 1
Print #1, TextBox.Text
Close 1
Exit Function
SaveFileError:
MsgBox "An error occurred with the textbox's 'SaveFile' command. Checked the path yet?", , "Error"
End Function

Public Function ParseString(Text As String)
On Error GoTo ParseError:
Dim BoxNumberF As Integer
Dim VarNumberF As Integer
Dim NewText As String

Dim MidStartInt As Integer
Dim MidEndInt As Integer

Dim ChrStr As Integer
Dim HexTempNum As Integer

NewText = Text

Do While InStr(1, NewText, "{Text(") <> 0
If InStr(1, NewText, "{Text(") <> 0 Then
BoxNumber = Mid(NewText, InStr(1, NewText, "{Text(") + 6)
BoxNumber = Mid(BoxNumber, 1, InStr(1, BoxNumber, ")") - 1)
BoxNumberF = BoxNumber
NewText = Replace(NewText, "{Text(" & BoxNumber & ")}", Form4.Text1(BoxNumberF).Text)
End If
Loop

Do While InStr(1, NewText, "{Var(") <> 0
If InStr(1, NewText, "{Var(") <> 0 Then
VarNumber = Mid(NewText, InStr(1, NewText, "{Var(") + 5)
VarNumber = Mid(VarNumber, 1, InStr(1, VarNumber, ")") - 1)
VarNumberF = VarNumber
NewText = Replace(NewText, "{Var(" & VarNumber & ")}", V.Vars.List(VarNumberF))
End If
Loop

Do While InStr(1, NewText, "{System.AppPath}") <> 0
NewText = Replace(NewText, "{System.AppPath}", App.Path)
Loop

Do While InStr(1, NewText, "{System.AppName}") <> 0
NewText = Replace(NewText, "{System.AppName}", App.EXEName)
Loop

Do While InStr(1, NewText, "{Asc(") <> 0
AscNum = Mid(NewText, InStr(1, NewText, "{Asc'") + 7)
AscNum = Mid(AscNum, 1, InStr(1, AscNum, "')}") - 1)

NewText = Replace(NewText, "{Asc('" & AscNum & "')}", Asc(AscNum))
Loop

Do While InStr(1, NewText, "{Asc(") <> 0
AscNum = Mid(NewText, InStr(1, NewText, "{Asc'") + 7)
AscNum = Mid(AscNum, 1, InStr(1, AscNum, "')}") - 1)

NewText = Replace(NewText, "{Asc('" & AscNum & "')}", Asc(AscNum))
Loop

Do While InStr(1, NewText, "{Chr(") <> 0
ChrNum = Mid(NewText, InStr(1, NewText, "{Chr(") + 5)
ChrNum = Mid(ChrNum, 1, InStr(1, ChrNum, ")}") - 1)
ChrStr = ChrNum

NewText = Replace(NewText, "{Chr(" & ChrNum & ")}", Chr(ChrStr))
Loop

Do While InStr(1, NewText, "{Hex(") <> 0
HexNum = Mid(NewText, InStr(1, NewText, "{Hex'") + 7)
HexNum = Mid(HexNum, 1, InStr(1, HexNum, "')}") - 1)

HexTempNum = Asc(HexNum)
HexChr = Hex(HexTempNum)

NewText = Replace(NewText, "{Hex('" & HexNum & "')}", HexChr)
Loop

'Mid() Function must be last so that it can use other sources.
Do While InStr(1, NewText, "{Mid(") <> 0
MidArg = Mid(NewText, InStr(1, NewText, "{Mid('"))
MidArg = Mid(NewText, 1, InStr(1, NewText, ")}") + 2)

MidString = Mid(MidArg, InStr(1, MidArg, "{Mid('") + 6)
MidString = Mid(MidString, 1, InStr(1, MidString, "',") - 1)

MidStartEnd = Mid(MidArg, InStr(1, MidArg, "',") + 2)
MidStartEnd = Mid(MidStartEnd, 1, InStr(1, MidStartEnd, ",") - 1)
MidStartInt = MidStartEnd

MidStartEnd = Mid(MidArg, InStr(1, MidArg, "',") + 2)
MidStartEnd = Mid(MidStartEnd, InStr(1, MidStartEnd, ",") + 1)
MidStartEnd = Mid(MidStartEnd, 1, InStr(1, MidStartEnd, ")}") - 1)
MidEndInt = MidStartEnd

NewText = Replace(NewText, MidArg, Mid(MidString, MidStartInt, MidEndInt))
Exit Do
Loop

ParseString = NewText
Exit Function
ParseError:
MsgBox "The string parsing function kicked the bucket, please check your variables. You should never really see this error so contact me please.", , "Error"
End Function

Public Function ExecuteAppCode(Code As String)
On Error GoTo ExecError:
Dim IfIf As Boolean
Dim IfElse As Boolean
Dim IfEndIf As Boolean

Dim CurNumberF As Integer
Dim CurValNum As Integer
Dim CurValStr As String
Dim CurValue As String
Dim MessageString As String
Dim VarNumberF As Integer
Dim CurIfValOne As String
Dim CurIfValTwo As String
Dim CurIfValOneNumF As Integer
Dim CurIfValTwoInt As Integer
Dim CurIfValTwoStr As String

CurCode = Split(Code, vbCrLf)
For CurLineOfCode = 0 To UBound(CurCode)

If Mid(CurCode(CurLineOfCode), 1, 4) = "Else" Then

If Mid(CurCode(CurLineOfCode), 5, 2) = "If" Then
CurCode(CurLineOfCode) = Mid(CurCode(CurLineOfCode), 5)
End If

    If IfIf = True Then
        IfEndIf = True
    ElseIf IfIf = False Then
        IfIf = False
        IfEndIf = False
    End If
IfElse = False
End If

If IfElse = True Then
GoTo EndLine:
End If

If Mid(CurCode(CurLineOfCode), 1, 6) = "End If" Then
IfEndIf = False
IfElse = False
IfIf = False
GoTo EndLine:
End If

If IfEndIf = True Then
GoTo EndLine:
End If

If Mid(CurCode(CurLineOfCode), 1, 7) = "MsgBox " Then
MessageString = Mid(CurCode(CurLineOfCode), 8)
MessageString = Replace(MessageString, Chr(34), "")
MessageString = ParseString(MessageString)
MsgBox MessageString
ElseIf Mid(CurCode(CurLineOfCode), 1, 10) = "CreateVar(" Then
VarNumber = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), "CreateVar(") + 10)
VarNumber = Mid(VarNumber, 1, InStr(1, VarNumber, ")") - 1)
VarNumberF = VarNumber
VarContent = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), "<-" & Chr(34)) + 3)
VarContent = Mid(VarContent, 1, InStr(1, VarContent, Chr(34)) - 1)
V.Vars.AddItem VarContent, VarNumberF
ElseIf Mid(CurCode(CurLineOfCode), 1, 9) = "ClearVars" Then
V.Vars.Clear
ElseIf Mid(CurCode(CurLineOfCode), 1, 4) = "Cmd(" Then
CurNumber = Mid(CurCode(CurLineOfCode), 5, InStr(1, CurCode(CurLineOfCode), ")") - 5)
CurNumberF = CurNumber
CurArgument = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), ").") + 2)
CurArgument = Mid(CurArgument, 1, InStr(1, CurArgument, " ") - 1)
If InStr(1, CurCode(CurLineOfCode), Chr(34)) = 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " "))
CurValue = Replace(CurValue, " = ", "")
CurValNum = CurValue
ElseIf InStr(1, CurCode(CurLineOfCode), Chr(34)) <> 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " = ") + 4)
CurValue = Mid(CurValue, 1, Len(CurValue) - 1)
CurValStr = ParseString(CurValue)
End If
If CurArgument = "Left" Then
Form4.Command1(CurNumberF).Left = CurValNum
ElseIf CurArgument = "Top" Then
Form4.Command1(CurNumberF).Top = CurValNum
ElseIf CurArgument = "Width" Then
Form4.Command1(CurNumberF).Width = CurValNum
ElseIf CurArgument = "Height" Then
Form4.Command1(CurNumberF).Height = CurValNum
ElseIf CurArgument = "Visible" Then
If CurValStr = "True" Then
Form4.Command1(CurNumberF).Visible = True
ElseIf CurValStr = "False" Then
Form4.Command1(CurNumberF).Visible = False
End If
ElseIf CurArgument = "Caption" Then
Form4.Command1(CurNumberF).Caption = CurValStr
End If
CurValStr = ""
CurValNum = 0
CurValue = ""
ElseIf Mid(CurCode(CurLineOfCode), 1, 5) = "Text(" Then
CurNumber = Mid(CurCode(CurLineOfCode), 6, InStr(1, CurCode(CurLineOfCode), ")") - 6)
CurNumberF = CurNumber
CurArgument = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), ").") + 2)
CurArgument = Mid(CurArgument, 1, InStr(1, CurArgument, " ") - 1)
If InStr(1, CurCode(CurLineOfCode), Chr(34)) = 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " "))
CurValue = Replace(CurValue, " = ", "")
CurValNum = CurValue
ElseIf InStr(1, CurCode(CurLineOfCode), Chr(34)) <> 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " = ") + 4)
CurValue = Mid(CurValue, 1, Len(CurValue) - 1)
CurValStr = ParseString(CurValue)
End If
If CurArgument = "Left" Then
Form4.Text1(CurNumberF).Left = CurValNum
ElseIf CurArgument = "Top" Then
Form4.Text1(CurNumberF).Top = CurValNum
ElseIf CurArgument = "Width" Then
Form4.Text1(CurNumberF).Width = CurValNum
ElseIf CurArgument = "Height" Then
Form4.Text1(CurNumberF).Height = CurValNum
ElseIf CurArgument = "Visible" Then
If CurValStr = "True" Then
Form4.Text1(CurNumberF).Visible = True
ElseIf CurValStr = "False" Then
Form4.Text1(CurNumberF).Visible = False
End If
ElseIf CurArgument = "Caption" Then
Form4.Text1(CurNumberF).Text = CurValStr
ElseIf CurArgument = "SaveFile" Then
SaveToFile Form4.Text1(CurNumberF), CurValStr
ElseIf CurArgument = "LoadFile" Then
LoadFromFile Form4.Text1(CurNumberF), CurValStr
End If
CurValStr = ""
CurValNum = 0
CurValue = ""
ElseIf Mid(CurCode(CurLineOfCode), 1, 6) = "Label(" Then
CurNumber = Mid(CurCode(CurLineOfCode), 7, InStr(1, CurCode(CurLineOfCode), ")") - 7)
CurNumberF = CurNumber
CurArgument = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), ").") + 2)
CurArgument = Mid(CurArgument, 1, InStr(1, CurArgument, " ") - 1)
If InStr(1, CurCode(CurLineOfCode), Chr(34)) = 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " "))
CurValue = Replace(CurValue, " = ", "")
CurValNum = CurValue
ElseIf InStr(1, CurCode(CurLineOfCode), Chr(34)) <> 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " = ") + 4)
CurValue = Mid(CurValue, 1, Len(CurValue) - 1)
CurValStr = ParseString(CurValue)
End If
If CurArgument = "Left" Then
Form4.Label1(CurNumberF).Left = CurValNum
ElseIf CurArgument = "Top" Then
Form4.Label1(CurNumberF).Top = CurValNum
ElseIf CurArgument = "Width" Then
Form4.Label1(CurNumberF).Width = CurValNum
ElseIf CurArgument = "Height" Then
Form4.Label1(CurNumberF).Height = CurValNum
ElseIf CurArgument = "Visible" Then
If CurValStr = "True" Then
Form4.Label1(CurNumberF).Visible = True
ElseIf CurValStr = "False" Then
Form4.Label1(CurNumberF).Visible = False
End If
ElseIf CurArgument = "Caption" Then
Form4.Label1(CurNumberF).Caption = CurValStr
End If
CurValStr = ""
CurValNum = 0
CurValue = ""
ElseIf Mid(CurCode(CurLineOfCode), 1, 6) = "Image(" Then
CurNumber = Mid(CurCode(CurLineOfCode), 7, InStr(1, CurCode(CurLineOfCode), ")") - 7)
CurNumberF = CurNumber
CurArgument = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), ").") + 2)
CurArgument = Mid(CurArgument, 1, InStr(1, CurArgument, " ") - 1)
If InStr(1, CurCode(CurLineOfCode), Chr(34)) = 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " "))
CurValue = Replace(CurValue, " = ", "")
CurValNum = CurValue
ElseIf InStr(1, CurCode(CurLineOfCode), Chr(34)) <> 0 Then
CurValue = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), " = ") + 4)
CurValue = Mid(CurValue, 1, Len(CurValue) - 1)
CurValStr = ParseString(CurValue)
End If
If CurArgument = "Left" Then
Form4.Image1(CurNumberF).Left = CurValNum
ElseIf CurArgument = "Top" Then
Form4.Image1(CurNumberF).Top = CurValNum
ElseIf CurArgument = "Width" Then
Form4.Image1(CurNumberF).Width = CurValNum
ElseIf CurArgument = "Height" Then
Form4.Image1(CurNumberF).Height = CurValNum
ElseIf CurArgument = "Visible" Then
If CurValStr = "True" Then
Form4.Image1(CurNumberF).Visible = True
ElseIf CurValStr = "False" Then
Form4.Image1(CurNumberF).Visible = False
End If
ElseIf CurArgument = "Picture" Then
Form4.Image1(CurNumberF).Picture = LoadPicture(CurValStr)
End If
CurValStr = ""
CurValNum = 0
CurValue = ""
ElseIf Mid(CurCode(CurLoneOfCode), 1, 5) = "Form." Then
CurArg = Mid(CurCode(CurLineOfCode), 6)
CurArg = Mid(CurArg, 1, InStr(1, CurArg, " ") - 1)

CurVal = Mid(CurCode(CurLineOfCode), 5 + Len(CurArg) + 4)

If InStr(1, CurVal, Chr(34)) <> 0 Then
CurValStr = Replace(CurVal, Chr(34), "")
CurValStr = ParseString(CurValStr)
Else
CurValNum = CurVal
End If

If CurArg = "Top" Then
Form4.Top = CurValNum
ElseIf CurArg = "Left" Then
Form4.Left = CurValNum
ElseIf CurArg = "Width" Then
Form4.Width = CurValNum
ElseIf CurArg = "Height" Then
Form4.Height = CurValNum
ElseIf CurArg = "Hide" Then
Form4.Hide
ElseIf CurArg = "Show" Then
Form4.Show
ElseIf CurArg = "Caption" Then
Form4.Caption = CurValStr
End If

ElseIf Mid(CurCode(CurLoneOfCode), 1, 7) = "System." Then
CurArg = Mid(CurCode(CurLineOfCode), 8)
CurArg = Mid(CurArg, 1, InStr(1, CurArg, " ") - 1)

CurVal = Mid(CurCode(CurLineOfCode), 9 + Len(CurArg))

If InStr(1, CurVal, Chr(34)) <> 0 Then
CurValStr = Replace(CurVal, Chr(34), "")
CurValStr = ParseString(CurValStr)
End If

If CurArg = "Copy" Then
On Error GoTo FileCopyErr:
CopyFileFrom = Mid(CurValStr, 1, InStr(1, CurValStr, "->") - 1)
CopyFileTo = Mid(CurValStr, InStr(1, CurValStr, "->") + 2)
FileCopy CopyFileFrom, CopyFileTo
FileCopyErr:
ElseIf CurArg = "Delete" Then
On Error GoTo FileDelErr:
Kill CurValStr
FileDelErr:
ElseIf CurArg = "Shell" Then
On Error GoTo FileShellErr
Shell CurValStr, vbNormalFocus
FileShellErr:
End If

ElseIf Mid(CurCode(CurLineOfCode), 1, 3) = "If " Then

'Get Case One Component
CurIfValOne = Mid(CurCode(CurLineOfCode), 4)
CurIfValOne = Mid(CurIfValOne, 1, InStr(1, CurIfValOne, " ") - 1)

'Get Case One Number
CurIfValOneNum = Replace(Mid(CurIfValOne, InStr(1, CurIfValOne, "(") + 1), ")", "")
CurIfValOneNum = Mid(CurIfValOneNum, 1, InStr(1, CurIfValOneNum, ".") - 1)
CurIfValOneNumF = CurIfValOneNum

'Get Case One Parameter
CurIfValOnePara = Mid(CurIfValOne, InStr(1, CurIfValOne, ".") + 1)

'Get Case One Component
CurIfValOne = Mid(CurCode(CurLineOfCode), 4)
CurIfValOne = Replace(Mid(CurIfValOne, 1, InStr(1, CurIfValOne, ".") - 1), "(", "")
CurIfValOne = Replace(CurIfValOne, ")", "")
CurIfValOne = Replace(CurIfValOne, CurIfValOneNum, "")

'Get Argument
CurIfArgument = Mid(CurCode(CurLineOfCode), Len(CurIfValOne) + 7 + Len(CurIfValOneNumF) + Len(CurIfValOnePara))
CurIfArgument = Mid(CurIfArgument, 1, InStr(1, CurIfArgument, " ") - 1)

'Get Case Two
CurIfValTwo = Mid(CurCode(CurLineOfCode), InStr(1, CurCode(CurLineOfCode), CurIfArgument) + 2)
CurIfValTwo = Mid(CurIfValTwo, 1, InStr(1, CurIfValTwo, " Then") - 1)
If InStr(1, CurIfValTwo, Chr(34)) <> 0 Then
CurIfValTwoStr = Replace(CurIfValTwo, Chr(34), "")
CurIfValTwoStr = ParseString(CurIfValTwoStr)
Else
CurIfValTwoInt = CurIfValTwo
End If
CurIfValTwo = ParseString(CurIfValTwo)

If CurIfValOne = "Cmd" Then

    If CurIfValOnePara = "Left" Then
        If CurIfArgument = "=" Then
            If Form4.Command1(CurIfValOneNumF).Left = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Command1(CurIfValOneNumF).Left > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Command1(CurIfValOneNumF).Left < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Command1(CurIfValOneNumF).Left <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CCurIfValOnePara = "Top" Then
        If CurIfArgument = "=" Then
            If Form4.Command1(CurIfValOneNumF).Top = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Command1(CurIfValOneNumF).Top > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Command1(CurIfValOneNumF).Top < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Command1(CurIfValOneNumF).Top <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Width" Then
        If CurIfArgument = "=" Then
            If Form4.Command1(CurIfValOneNumF).Width = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Command1(CurIfValOneNumF).Width > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Command1(CurIfValOneNumF).Width < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Command1(CurIfValOneNumF).Width <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Height" Then
        If CurIfArgument = "=" Then
            If Form4.Command1(CurIfValOneNumF).Height = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Command1(CurIfValOneNumF).Height > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Command1(CurIfValOneNumF).Height < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Command1(CurIfValOneNumF).Height <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Visible" Then
        If CurIfArgument = "=" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Command1(CurIfValOneNumF).Visible = True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Command1(CurIfValOneNumF).Visible = False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        
        ElseIf CurIfArgument = "<>" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Command1(CurIfValOneNumF).Visible <> True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Command1(CurIfValOneNumF).Visible <> False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        End If
    ElseIf CurIfValOnePara = "Caption" Then
        If CurIfArgument = "=" Then
            If Form4.Command1(CurIfValOneNumF).Caption = CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Command1(CurIfValOneNumF).Height <> CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    End If

ElseIf CurIfValOne = "Text" Then

    If CurIfValOnePara = "Left" Then
        If CurIfArgument = "=" Then
            If Form4.Text1(CurIfValOneNumF).Left = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Text1(CurIfValOneNumF).Left > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Text1(CurIfValOneNumF).Left < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Text1(CurIfValOneNumF).Left <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CCurIfValOnePara = "Top" Then
        If CurIfArgument = "=" Then
            If Form4.Text1(CurIfValOneNumF).Top = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Text1(CurIfValOneNumF).Top > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Text1(CurIfValOneNumF).Top < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Text1(CurIfValOneNumF).Top <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Width" Then
        If CurIfArgument = "=" Then
            If Form4.Text1(CurIfValOneNumF).Width = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Text1(CurIfValOneNumF).Width > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Text1(CurIfValOneNumF).Width < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Text1(CurIfValOneNumF).Width <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Height" Then
        If CurIfArgument = "=" Then
            If Form4.Text1(CurIfValOneNumF).Height = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Text1(CurIfValOneNumF).Height > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Text1(CurIfValOneNumF).Height < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Text1(CurIfValOneNumF).Height <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Visible" Then
        If CurIfArgument = "=" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Text1(CurIfValOneNumF).Visible = True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Text1(CurIfValOneNumF).Visible = False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        
        ElseIf CurIfArgument = "<>" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Text1(CurIfValOneNumF).Visible <> True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Text1(CurIfValOneNumF).Visible <> False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        End If
    ElseIf CurIfValOnePara = "Caption" Then
        If CurIfArgument = "=" Then
            If Form4.Text1(CurIfValOneNumF).Text = CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Text1(CurIfValOneNumF).Height <> CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    End If

ElseIf CurIfValOne = "Label" Then


    If CurIfValOnePara = "Left" Then
        If CurIfArgument = "=" Then
            If Form4.Label1(CurIfValOneNumF).Left = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Label1(CurIfValOneNumF).Left > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Label1(CurIfValOneNumF).Left < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Label1(CurIfValOneNumF).Left <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CCurIfValOnePara = "Top" Then
        If CurIfArgument = "=" Then
            If Form4.Label1(CurIfValOneNumF).Top = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Label1(CurIfValOneNumF).Top > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Label1(CurIfValOneNumF).Top < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Label1(CurIfValOneNumF).Top <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Width" Then
        If CurIfArgument = "=" Then
            If Form4.Label1(CurIfValOneNumF).Width = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Label1(CurIfValOneNumF).Width > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Label1(CurIfValOneNumF).Width < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Label1(CurIfValOneNumF).Width <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Height" Then
        If CurIfArgument = "=" Then
            If Form4.Label1(CurIfValOneNumF).Height = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Label1(CurIfValOneNumF).Height > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Label1(CurIfValOneNumF).Height < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Label1(CurIfValOneNumF).Height <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Visible" Then
        If CurIfArgument = "=" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Label1(CurIfValOneNumF).Visible = True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Label1(CurIfValOneNumF).Visible = False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        
        ElseIf CurIfArgument = "<>" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Label1(CurIfValOneNumF).Visible <> True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Label1(CurIfValOneNumF).Visible <> False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        End If
    ElseIf CurIfValOnePara = "Caption" Then
        If CurIfArgument = "=" Then
            If Form4.Label1(CurIfValOneNumF).Caption = CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Label1(CurIfValOneNumF).Height <> CurIfValTwoStr Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    End If


ElseIf CurIfValOne = "Image" Then

    If CurIfValOnePara = "Left" Then
        If CurIfArgument = "=" Then
            If Form4.Image1(CurIfValOneNumF).Left = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Image1(CurIfValOneNumF).Left > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Image1(CurIfValOneNumF).Left < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Image1(CurIfValOneNumF).Left <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CCurIfValOnePara = "Top" Then
        If CurIfArgument = "=" Then
            If Form4.Image1(CurIfValOneNumF).Top = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Image1(CurIfValOneNumF).Top > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Image1(CurIfValOneNumF).Top < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Image1(CurIfValOneNumF).Top <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Width" Then
        If CurIfArgument = "=" Then
            If Form4.Image1(CurIfValOneNumF).Width = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Image1(CurIfValOneNumF).Width > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Image1(CurIfValOneNumF).Width < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Image1(CurIfValOneNumF).Width <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Height" Then
        If CurIfArgument = "=" Then
            If Form4.Image1(CurIfValOneNumF).Height = CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = ">" Then
            If Form4.Image1(CurIfValOneNumF).Height > CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<" Then
            If Form4.Image1(CurIfValOneNumF).Height < CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfArgument = "<>" Then
            If Form4.Image1(CurIfValOneNumF).Height <> CurIfValTwoInt Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
    ElseIf CurIfValOnePara = "Visible" Then
        If CurIfArgument = "=" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Image1(CurIfValOneNumF).Visible = True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Image1(CurIfValOneNumF).Visible = False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        
        ElseIf CurIfArgument = "<>" Then
        If CurIfValTwoStr = "True" Then
            If Form4.Image1(CurIfValOneNumF).Visible <> True Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        ElseIf CurIfValTwoStr = "False" Then
            If Form4.Image1(CurIfValOneNumF).Visible <> False Then
                IfElse = False
                IfIf = True
                GoTo EndLine:
            Else
                IfElse = True
                GoTo EndLine:
            End If
        End If
        End If
    End If
End If
























EndLine:
'If IfElse = True Then
'ElseIf IfElse = False Then
'End If
End If
Next CurLineOfCode

IfIf = False
IfElse = False
IfEndIf = False

CurNumberF = 0
CurValNum = 0
CurValStr = ""
CurValue = ""
MessageString = ""
VarNumberF = 0
CurIfValOne = ""
CurIfValTwo = ""
CurIfValOneNumF = 0
CurIfValTwoInt = 0
CurIfValTwoStr = ""
Exit Function
ExecError:
MsgBox "Your code/syntax/malicious use of the language has caused a horrible bug, please try to fix, or save the files and send them to me to fix.", , "Error"
End Function
