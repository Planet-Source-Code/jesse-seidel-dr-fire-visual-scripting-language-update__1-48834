VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form3 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Application Code"
   ClientHeight    =   3990
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   5760
   Icon            =   "AppCodeDevSDK.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VectorBasic.chameleonButton chameleonButton1 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      ToolTipText     =   "Start"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "AppCodeDevSDK.frx":030A
      PICN            =   "AppCodeDevSDK.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Vars"
      Height          =   255
      Left            =   4560
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Build"
      Height          =   255
      Left            =   5160
      TabIndex        =   4
      Top             =   2640
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   3000
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   2055
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
   Begin VectorBasic.chameleonButton chameleonButton2 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      ToolTipText     =   "Stop"
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "AppCodeDevSDK.frx":05B4
      PICN            =   "AppCodeDevSDK.frx":05D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Environment Code:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Application Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Code File"
      End
      Begin VB.Menu mnuFileBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Code File"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Code File"
      End
      Begin VB.Menu mnuFileBreakTwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCompile 
         Caption         =   "Compile Application"
      End
      Begin VB.Menu mnuFileBreakThree 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRun 
      Caption         =   "Run"
      Begin VB.Menu mnuRunStart 
         Caption         =   "Start"
      End
      Begin VB.Menu mnuRunStop 
         Caption         =   "Stop"
      End
   End
   Begin VB.Menu mnuEnvironment 
      Caption         =   "Environment"
      Begin VB.Menu mnuEnvironmentClear 
         Caption         =   "Clear Environment Code"
      End
      Begin VB.Menu mnuEnvironmentBreak 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnvironmentOpen 
         Caption         =   "Open Environment Code File"
      End
      Begin VB.Menu mnuEnvironmentSave 
         Caption         =   "Save Environment Code File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpTutorials 
         Caption         =   "Tutorials"
      End
      Begin VB.Menu mnuHelpBreakOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpSyntax 
         Caption         =   "Syntax Help"
      End
      Begin VB.Menu mnuHelpAppHelp 
         Caption         =   "SDK Help"
      End
      Begin VB.Menu mnuHelpBreakTwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim MouseDown As Integer
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2

Private Declare Function GetWindowLong Lib "user32" _
  Alias "GetWindowLongA" (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
   Alias "SetWindowLongA" (ByVal hwnd As Long, _
   ByVal nIndex As Long, ByVal dwNewLong As Long) _
   As Long

Private Declare Function SetLayeredWindowAttributes Lib _
    "user32" (ByVal hwnd As Long, ByVal crKey As Long, _
    ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Dim oOld As Long
Dim oNew As Long
Dim aOld As Long
Dim aNew As Long
Dim i As Long
Dim Incoming As Long
Dim Outgoing As Long
Dim temp1 As Long
Dim r1 As Single
Dim r2 As Single

Dim X As Long
Dim X1 As Long
Dim X2 As Long
Dim x3 As Long
Dim startPause As Boolean

Dim tValue As Long
Dim aValue As Long
Dim A As Integer
Dim OldX As Long, OldY As Long, IsMoving As Boolean
Dim Selected As Integer, Stuffin As Boolean
Dim ChangingCaption As Boolean

Dim EnvironmentCode As String
Dim ApplicationPath As String
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
Dim CurPicture As String
Dim CurStretch As String
Dim CurControlSpecs As String
Dim CurEXEFile As String
Dim EndOfFile As String
Dim CurLineFromFile As String
Private Sub chameleonButton1_Click()
On Error GoTo RunError:
Command1_Click
If Len(Text2.Text) < 10 Then
MsgBox "No Environment Code found! The Program was not executed!", , "Error"
Exit Sub
End If
Form1.Hide
Form2.Hide
Form3.Height = 975
Form4.Show
V.Show
V.Left = Me.Left + 5250
V.Top = Me.Top
Exit Sub
RunError:
MsgBox "An error has occurred, the program could not be run!", , "Error"
End Sub

Private Sub chameleonButton2_Click()
On Error GoTo StopRunError:
Form1.Show
Form2.Show
Form3.Height = 4275
Form4.Hide
Unload Form4
V.Vars.Clear
Exit Sub
StopRunError:
MsgBox "An error has occurred, the program could not be stopped!", , "Error"
End Sub

Private Sub Command1_Click()
Text2.Text = ""
On Error GoTo BuildingError:
BuildEnvironmentFormCode Form1
BuildEnvironmentCode Form1.Command1, "Command1", Form1
BuildEnvironmentCode Form1.Text1, "Text1", Form1
BuildEnvironmentCode Form1.Label1, "Label1", Form1
BuildEnvironmentCode Form1.Image1, "Image1", Form1
Exit Sub
BuildingError:
MsgBox "The environment code could not be built, try again, then contact me if the error persists!", , "Error"
End Sub

Private Sub Command2_Click()
V.Show
V.Left = Me.Left + 5250
V.Top = Me.Top
End Sub

Private Sub mnuEnvironmentClear_Click()
Text2.Text = ""
End Sub

Private Sub mnuEnvironmentOpen_Click()
On Error GoTo EnvOpenFileError:

Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Environment Code File|*.ibf"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Input As 1
Text2.Text = ""
Do Until EOF(1)
'Line Input #1, CurLineFromFile
'Text2.Text = Text2.Text & CurLineFromFile & vbCrLf
Loop
Close 1

'Clear Form
Unload Form1
Form1.Show

EnvironmentCode = Form3.Text2.Text

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
formwidthf = Mid(FormAnalysis, InStr(1, FormAnalysis, "Width:") + 6)
formwidthf = Mid(formwidthf, 1, InStr(1, formwidthf, " ") - 1)
formwidthf = formwidthf
FormHeightF = Mid(FormAnalysis, InStr(1, FormAnalysis, "Height:") + 7)
FormHeightF = Mid(FormHeightF, 1, InStr(1, FormHeightF, " ") - 1)
FormHeightF = FormHeightF
FormCaption = Mid(FormAnalysis, InStr(1, FormAnalysis, "Caption:") + 8)
FormCaption = Mid(FormCaption, 1, InStr(1, FormCaption, Chr(13)) - 1)

'Construct Form
Form1.Width = formwidthf
Form1.Height = FormHeightF
Form1.Caption = FormCaption
End If

Do While Len(EnvironmentCode) > 2
CurControlSpecs = Mid(EnvironmentCode, 1, InStr(1, EnvironmentCode, vbCrLf))
EnvironmentCode = Mid(EnvironmentCode, InStr(1, EnvironmentCode, vbCrLf) + 2)
CurControlSpecs = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Create Ctrl ") + 12)

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
CurLeftF = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Left:") + 5)
CurLeftF = Mid(CurLeftF, 1, InStr(1, CurLeftF, " ") - 1)
CurLeftF = CurLeftF

'Top
CurTopF = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Top:") + 4)
CurTopF = Mid(CurTopF, 1, InStr(1, CurTopF, " ") - 1)
CurTopF = CurTopF

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

'Stretch
CurStretch = Mid(CurControlSpecs, InStr(1, CurControlSpecs, "Stretch") + 8)
CurStretch = Mid(CurStretch, 1, InStr(1, CurStretch, Chr(13)) - 1)
End If

If CurName = "Cmd" Then
    If CurVisible = "True" Then
CreateCommandButton Form1.Command1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateCommandButton Form1.Command1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Text" Then
    If CurVisible = "True" Then
CreateTextBox Form1.Text1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateTextBox Form1.Text1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Label" Then
    If CurVisible = "True" Then
CreateLabel Form1.Label1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurCaption
    ElseIf CurVisible = "False" Then
CreateLabel Form1.Label1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurCaption
    End If
ElseIf CurName = "Image" Then
    If CurVisible = "True" Then
        If CurStretch = "True" Then
CreateImageBox Form1.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurPicture, True
        ElseIf CurStretch = "False" Then
CreateImageBox Form1.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, True, CurPicture, False
        End If
    ElseIf CurVisible = "False" Then
        If CurStretch = "True" Then
CreateImageBox Form1.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurPicture, True
        ElseIf CurStretch = "False" Then
CreateImageBox Form1.Image1, CurLeftF, CurTopF, CurWidthF, CurHeightF, False, CurPicture, False
        End If
    End If
End If

Loop
Exit Sub
EnvOpenFileError:
'MsgBox "The environment code was not valid. Or the path was wrong, try again!", , "Error"
End Sub

Private Sub mnuEnvironmentSave_Click()
On Error GoTo EnvSaveFileError:
Command1_Click
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Environment Code File|*.ibf"
CommonDialog1.ShowSave
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Output As 1
Print #1, Text2.Text
Close 1
Exit Sub
EnvSaveFileError:
'MsgBox "The code couldn't be saved, check the path.", , "Error"
End Sub

Private Sub mnuFileCompile_Click()
On Error GoTo CompileError:
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Win32 Executable File|*.exe"
CommonDialog1.ShowOpen
CurEXEFile = CommonDialog1.FileName
If Right(App.Path, 1) = "\" Then
ApplicationPath = App.Path
ElseIf Right(App.Path, 1) <> "\" Then
ApplicationPath = App.Path & "\"
End If
FileCopy ApplicationPath & "Exec.dll", CurEXEFile
Open CurEXEFile For Binary As #1
EndOfFile = LOF(1)
Put #1, EndOfFile, "(<EnvironmentEXECode>)" & Text2.Text & "(<ApplicationEXECode>)" & Text1.Text
Close #1
Exit Sub
CompileError:
If Err.Number = 53 Then
MsgBox "The EXE could not compile. Check 'Exec.dll' is in the same folder as the IDE.", , "Error"
Err.Number = 0
End If
End Sub

Private Sub mnuFileExit_Click()
Unload Form1
Unload Form2
Unload Form4
Unload V
Unload Me
End
End Sub

Private Sub mnuFileNew_Click()
Text1.Text = ""
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo OpenCodeError:
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Application Code File|*.ibc"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Input As 1
Text1.Text = ""
Do Until EOF(1)
Line Input #1, CurLineFromFile
Text1.Text = Text1.Text & CurLineFromFile & vbCrLf
Loop
Close 1
Exit Sub
OpenCodeError:
'MsgBox "An error has occurred whilst trying to open a code file. Please check the path!", , "Error"
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo SaveCodeError:
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Application Code File|*.ibc"
CommonDialog1.ShowSave
CurEnvFile = CommonDialog1.FileName
Open CurEnvFile For Output As 1
Print #1, Text1.Text
Close 1
Exit Sub
SaveCodeError:
'MsgBox "An error has occurred whilst trying to save a code file. Please check the path!", , "Error"
End Sub

Private Sub mnuHelpAbout_Click()
MsgBox "Inter Basic" & vbCrLf & "Made by Jesse Seidel of Inter3" & vbCrLf & "Visit: Www.Inter3.NET", vbOKOnly, "About"
End Sub

Private Sub mnuHelpSyntax_Click()
CommonDialog1.HelpFile = App.HelpFile
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub

Private Sub mnuHelpTutorials_Click()
CommonDialog1.HelpFile = App.HelpFile
CommonDialog1.HelpCommand = cdlHelpContents
CommonDialog1.ShowHelp
End Sub

Private Sub mnuRunStart_Click()
On Error GoTo RunError:
Command1_Click
If Len(Text2.Text) < 10 Then
MsgBox "No Environment Code found! The Program was not executed!", , "Error"
Exit Sub
End If
Form1.Hide
Form2.Hide
Form3.Height = 645
Form4.Show
V.Show
V.Left = Me.Left + 5250
V.Top = Me.Top
Exit Sub
RunError:
MsgBox "An error has occurred, the program could not be run!", , "Error"
End Sub

Private Sub mnuRunStop_Click()
On Error GoTo StopRunError:
Form1.Show
Form2.Show
Form3.Height = 4275
Form4.Hide
Unload Form4
V.Vars.Clear
Exit Sub
StopRunError:
MsgBox "An error has occurred, the program could not be stopped!", , "Error"
End Sub
