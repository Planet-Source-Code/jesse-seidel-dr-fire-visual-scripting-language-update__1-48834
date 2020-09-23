VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form V 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Variables"
   ClientHeight    =   2850
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   3030
   Icon            =   "Vars.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2520
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox Vars 
      Height          =   2595
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Variables:"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileClear 
         Caption         =   "Clear Variables"
      End
      Begin VB.Menu mnuFileBreakOne 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Variables"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save Variables"
      End
      Begin VB.Menu mnuFileBreakTwo 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "V"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuFileClear_Click()
Vars.Clear
End Sub

Private Sub mnuFileExit_Click()
Me.Hide
End Sub

Private Sub mnuFileOpen_Click()
On Error GoTo OpenVarError:
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Variables Storage File|*.ibv"
CommonDialog1.ShowOpen
CurEnvFile = CommonDialog1.FileName

Open CurEnvFile For Input As 1
Do Until EOF(1)
Line Input #1, CurLineFromFile
Vars.AddItem CurLineFromFile
Loop
Close 1
Exit Sub
OpenVarError:
MsgBox "The variables were not saved, checked the path and content?", , "Error"
End Sub

Private Sub mnuFileSave_Click()
On Error GoTo SaveVarError:
Dim CurEnvFile As String
CommonDialog1.FileName = ""
CommonDialog1.Filter = "InterBasic Variables Storage File|*.ibv"
CommonDialog1.ShowSave
CurEnvFile = CommonDialog1.FileName

Open CurEnvFile For Output As 1
For DoMeForLong = 1 To Vars.ListCount
Print #1, Vars.List(DoMeForLong - 1)
Next DoMeForLong
Close 1
Exit Sub
SaveVarError:
MsgBox "The variables were not saved, checked the path and content?", , "Error"
End Sub
