VERSION 5.00
Begin VB.UserControl UserControl1 
   BackStyle       =   0  'Transparent
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   MousePointer    =   5  'Size
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   100
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   100
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000010&
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   3375
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'PLEASE VOTE as this usercontrol is the result of a lot of work
'Bryan Cairns - cairnsb@html-helper.com
'
'All you need to do to "activate" the "grippers" is to set the "BoundControl" property
'the usercontrol will take case of the rest


Dim m_MoveMode As Boolean
Dim m_BoundControl As Object
Dim XX As Single
Dim YY As Single

Dim SX As Single
Dim SY As Single

Const m_def_MoveMode = 0

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Resize()
Event Moving()


Private Sub ChangeMoveState()
    If m_MoveMode = False Then
    m_MoveMode = True
    Else
    m_MoveMode = False
    End If
    Shape1.Visible = m_MoveMode
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SX = X
    YY = Y
    If Button = 2 Then ChangeMoveState
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then ResizeBoundControl Index, X, Y
End Sub

Private Sub UserControl_Initialize()

Label1(0).Move 0, 0, 100, 100

Dim i As Integer
MoveGrips 0
For i = 1 To 7
    Load Label1(i)
    MoveGrips i
    Label1(i).Visible = True
     Label1(i).ZOrder 0
Next i
SetControlOnTop
End Sub

Private Sub MoveGrips(i As Integer)
On Error Resume Next
Select Case i
Case Is = 0 'top left
    Label1(i).Move 0, 0
    Label1(i).MousePointer = 8
Case Is = 1 'top middle
    Label1(i).Move (UserControl.Width / 2) - (Label1(i).Width / 2), 0
    Label1(i).MousePointer = 7
Case Is = 2 'top right
    Label1(i).Move UserControl.Width - Label1(i).Width, 0
    Label1(i).MousePointer = 6
Case Is = 3 'middle left
    Label1(i).Move 0, (UserControl.Height / 2) - (Label1(i).Height / 2)
    Label1(i).MousePointer = 9
Case Is = 4 'middle right
    Label1(i).Move UserControl.Width - Label1(i).Width, (UserControl.Height / 2) - (Label1(i).Height / 2)
    Label1(i).MousePointer = 9
Case Is = 5 'bottom left
    Label1(i).Move 0, UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 6
Case Is = 6 'bottom middle
    Label1(i).Move (UserControl.Width / 2) - (Label1(i).Width / 2), UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 7
Case Is = 7 ' bottom right
    Label1(i).Move UserControl.Width - Label1(i).Width, UserControl.Height - Label1(i).Height
    Label1(i).MousePointer = 8
End Select
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    XX = X
    YY = Y
    If Button = 1 Then MoveBoundControl X, Y
    If Button = 2 Then ChangeMoveState
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then MoveBoundControl X, Y
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
Dim i As Integer
For i = 1 To 8
    MoveGrips i
Next i
Shape1.Move Label1(0).Width, Label1(0).Height, UserControl.Width - (Label1(0).Width * 2), UserControl.Height - (Label1(0).Height * 2)
End Sub

Public Property Get BoundControl() As Object
    Set BoundControl = m_BoundControl
End Property

Public Property Set BoundControl(ByVal New_BoundControl As Object)
    
    Set m_BoundControl = New_BoundControl
    PropertyChanged "BoundControl"

    If New_BoundControl Is Nothing Then
    
    Else
    SetSameContainer
    MoveUserControl
    End If
    
End Property

Private Sub SetSameContainer()
'Make sure the usercontrol is in the same container as the bound control
Dim i As Long
Dim OBJ As Object
On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(i).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i
If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub
    If m_BoundControl.Container <> OBJ.Container Then
        Set OBJ.Container = m_BoundControl.Container
    End If
End Sub
Private Sub MoveUserControl()
'find this control and move it on the parent container
Dim i As Long
Dim OBJ As Object
On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(i).hwnd Then
    Set OBJ = UserControl.ParentControls.Item(i)
        OBJ.Move m_BoundControl.Left - offset, m_BoundControl.Top - offset, m_BoundControl.Width + (offset * 2), m_BoundControl.Height + (offset * 2)
        Exit For
    End If
Next i

End Sub

Private Sub SetControlOnTop()
Dim i As Long
Dim OBJ As Object

On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1
    'get this usercontrol from the parent
    If UserControl.hwnd = UserControl.ParentControls.Item(i).hwnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        OBJ.ZOrder 0
        Exit For
    End If
Next i


End Sub
Private Sub ResizeBoundControl(ID As Integer, X, Y)
Dim i As Long
Dim OBJ As Object

On Error Resume Next
For i = 0 To UserControl.ParentControls.Count - 1
    'get this usercontrol from the parent
    If UserControl.hwnd = UserControl.ParentControls.Item(i).hwnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i

If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub

'figure out which "handle" they used to resize it with and resize the bound control

Select Case ID
Case Is = 0 'top left
    OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
Case Is = 1 'top middle
    OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
Case Is = 2 'top right
    OBJ.Top = OBJ.Top + Y
    OBJ.Height = OBJ.Height - Y
    OBJ.Width = OBJ.Width + (X)
Case Is = 3 'middle left
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
Case Is = 4 'middle right
    OBJ.Width = OBJ.Width + (X)
Case Is = 5 'bottom left
    OBJ.Left = OBJ.Left + X
    OBJ.Width = OBJ.Width - X
    OBJ.Height = OBJ.Height + (Y)
Case Is = 6 'bottom middle
    OBJ.Height = OBJ.Height + (Y)
Case Is = 7 ' bottom right
    OBJ.Height = OBJ.Height + (Y)
    OBJ.Width = OBJ.Width + (X)
End Select
m_BoundControl.Move OBJ.Left + offset, OBJ.Top + offset, OBJ.Width - (offset * 2), OBJ.Height - (offset * 2)

'make sure the usercontrol is the same size as the bound control
'This is needed as some controls (like text boxes) have a min width and height
If OBJ.Width < m_BoundControl.Width Or OBJ.Height < m_BoundControl.Height Then
    OBJ.Move m_BoundControl.Left - offset, m_BoundControl.Top - offset, m_BoundControl.Width + (offset * 2), m_BoundControl.Height + (offset * 2)
End If
End Sub

Private Sub MoveBoundControl(X As Single, Y As Single)
Dim i As Long
Dim OBJ As Object
On Error Resume Next
If m_MoveMode = False Then Exit Sub
For i = 0 To UserControl.ParentControls.Count - 1

    If UserControl.hwnd = UserControl.ParentControls.Item(i).hwnd Then
        Set OBJ = UserControl.ParentControls.Item(i)
        Exit For
    End If
Next i

If m_BoundControl Is Nothing Then Exit Sub
If OBJ Is Nothing Then Exit Sub

    OBJ.Move OBJ.Left + (X - XX), OBJ.Top + (Y - YY)
    m_BoundControl.Move m_BoundControl.Left + (X - XX), m_BoundControl.Top + (Y - YY)

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set m_BoundControl = PropBag.ReadProperty("BoundControl", Nothing)
    m_MoveMode = PropBag.ReadProperty("MoveMode", m_def_MoveMode)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BoundControl", m_BoundControl, Nothing)
    Call PropBag.WriteProperty("MoveMode", m_MoveMode, m_def_MoveMode)
End Sub

Public Property Get offset() As Single
    offset = Label1(0).Width
End Property

Public Property Get MoveMode() As Boolean
    MoveMode = m_MoveMode
End Property

Public Property Let MoveMode(ByVal New_MoveMode As Boolean)
    m_MoveMode = New_MoveMode
    PropertyChanged "MoveMode"
End Property

Private Sub UserControl_InitProperties()
    m_MoveMode = m_def_MoveMode
End Sub

