VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "New Application"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5130
   DrawWidth       =   3
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2160
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VectorBasic.chameleonButton Command1 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Button"
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
      MICON           =   "Form1.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VectorBasic.UserControl1 resize1 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin VB.TextBox PicPath 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   3600
      TabIndex        =   1
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
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
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
'
Private Sub Command1_Click(Index As Integer)
If InStr(1, Form3.Text1.Text, "Start Cmd(" & Index & ")_Clicked()") = 0 Then
Form3.Text1.Text = Form3.Text1.Text & "Start Cmd(" & Index & ")_Clicked()" & vbCrLf & vbCrLf & "End Cmd(" & Index & ")_Clicked" & vbCrLf & vbCrLf
End If

resize1.Visible = True
    Set resize1.BoundControl = Command1(Index)
End Sub


Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X: OldY = Y: IsMoving = True

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsMoving Then
                Command1(Index).Top = Command1(Index).Top - (OldY - Y)
                Command1(Index).Left = Command1(Index).Left - (OldX - X)
                End If
End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
IsMoving = False

End Sub

Private Sub Form_Activate()
Set selectedForm = Me
End Sub

Private Sub Form_Click()
resize1.Visible = False
End Sub


Private Sub Form_Load()
resize1.Visible = False
Dim OBJ As Object
DrawTheGrid Me
Form6.Visible = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
Form5.Label1.Caption = "'X' Position: " & X
Form5.Label2.Caption = "'Y' Position: " & Y
End If
End Sub

Private Sub Form_Resize()
Form2.Text6 = Me.Width
Form2.Text7 = Me.Height
DrawTheGrid Me
End Sub

Private Sub Image1_Click(Index As Integer)
MsgBox "My Control Name is: " & "Image(" & Index & ")"
resize1.Visible = True
    Set resize1.BoundControl = Image1(Index)
End Sub


Private Sub Label1_Click(Index As Integer)
MsgBox "My Control Name is: " & "Label(" & Index & ")"
resize1.Visible = True
    Set resize1.BoundControl = Label1(Index)
End Sub


Private Sub PicPath_Click(Index As Integer)
resize1.Visible = True
    Set resize1.BoundControl = PicPath(Index)
End Sub


Private Sub Text1_Click(Index As Integer)
MsgBox "My Control Name is: " & "Text(" & Index & ")"
resize1.Visible = True
Set resize1.BoundControl = Text1(Index)
End Sub

