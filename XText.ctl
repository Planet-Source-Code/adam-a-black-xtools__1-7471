VERSION 5.00
Begin VB.UserControl XText 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1530
   ScaleHeight     =   315
   ScaleWidth      =   1530
   ToolboxBitmap   =   "XText.ctx":0000
   Begin VB.TextBox txtParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   290
      IMEMode         =   3  'DISABLE
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "XText.ctx":0312
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "XText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Dim FocusState As Boolean
Dim m_FillColor As OLE_COLOR
Dim m_FocusColor As OLE_COLOR
Const m_def_FillColor = &HFFE3B5
Const m_def_FocusColor = vbWhite
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Public Property Get ForeColor() As OLE_COLOR
ForeColor = txtParent.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
txtParent.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
Enabled = txtParent.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
txtParent.Enabled() = New_Enabled
PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Set Font = txtParent.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
Set txtParent.Font = New_Font
PropertyChanged "Font"
End Property
Public Property Get BorderStyle() As Integer
BorderStyle = txtParent.BorderStyle
End Property
Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
txtParent.BorderStyle() = New_BorderStyle
PropertyChanged "BorderStyle"
End Property
Public Sub Refresh()
txtParent.Refresh
End Sub
Private Sub txtParent_Click()
RaiseEvent Click
End Sub
Private Sub txtParent_DblClick()
RaiseEvent DblClick
End Sub
Private Sub txtParent_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtParent_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub txtParent_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub txtParent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub txtParent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
If FocusState = False Then
With txtParent
If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
ReleaseCapture
txtParent.BackColor = FillColor
Else
SetCapture .hWnd
txtParent.BackColor = FocusColor
End If
End With
End If
End Sub
Private Sub txtParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Public Property Get Alignment() As Integer
Alignment = txtParent.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
txtParent.Alignment() = New_Alignment
PropertyChanged "Alignment"
End Property
Private Sub txtParent_Change()
RaiseEvent Change
End Sub
Public Property Get Locked() As Boolean
Locked = txtParent.Locked
End Property
Public Property Let Locked(ByVal New_Locked As Boolean)
txtParent.Locked() = New_Locked
PropertyChanged "Locked"
End Property
Public Property Get MaxLength() As Long
MaxLength = txtParent.MaxLength
End Property
Public Property Let MaxLength(ByVal New_MaxLength As Long)
txtParent.MaxLength() = New_MaxLength
PropertyChanged "MaxLength"
End Property
Public Property Get Text() As String
Text = txtParent.Text
End Property
Public Property Let Text(ByVal New_Text As String)
txtParent.Text() = New_Text
PropertyChanged "Text"
End Property
Public Property Get FocusColor() As OLE_COLOR
FocusColor = m_FocusColor
End Property
Public Property Let FocusColor(ByVal New_FocusColor As OLE_COLOR)
m_FocusColor = New_FocusColor
PropertyChanged "FocusColor"
End Property
Private Sub UserControl_EnterFocus()
FocusState = True
If FocusState = True Then
txtParent.BackColor = FocusColor
End If
End Sub
Private Sub UserControl_ExitFocus()
FocusState = False
If FocusState = False Then
txtParent.BackColor = FillColor
End If
End Sub
Private Sub UserControl_InitProperties()
txtParent.Text = Extender.Name
m_FocusColor = m_def_FocusColor
m_FillColor = m_def_FillColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
txtParent.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
txtParent.Enabled = PropBag.ReadProperty("Enabled", True)
Set txtParent.Font = PropBag.ReadProperty("Font", Ambient.Font)
txtParent.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
txtParent.Alignment = PropBag.ReadProperty("Alignment", 0)
txtParent.Locked = PropBag.ReadProperty("Locked", False)
txtParent.MaxLength = PropBag.ReadProperty("MaxLength", 0)
txtParent.Text = PropBag.ReadProperty("Text", "XText")
m_FocusColor = PropBag.ReadProperty("FocusColor", m_def_FocusColor)
m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
txtParent.BackColor = FillColor
End Sub
Private Sub UserControl_Resize()
txtParent.Height = UserControl.Height
txtParent.Width = UserControl.Width
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("ForeColor", txtParent.ForeColor, &H80000008)
Call PropBag.WriteProperty("Enabled", txtParent.Enabled, True)
Call PropBag.WriteProperty("Font", txtParent.Font, Ambient.Font)
Call PropBag.WriteProperty("BorderStyle", txtParent.BorderStyle, 1)
Call PropBag.WriteProperty("Alignment", txtParent.Alignment, 0)
Call PropBag.WriteProperty("Locked", txtParent.Locked, False)
Call PropBag.WriteProperty("MaxLength", txtParent.MaxLength, 0)
Call PropBag.WriteProperty("Text", txtParent.Text, "XText")
Call PropBag.WriteProperty("FocusColor", m_FocusColor, m_def_FocusColor)
Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
End Sub
Public Property Get FillColor() As OLE_COLOR
FillColor = m_FillColor
txtParent.BackColor = FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
m_FillColor = New_FillColor
PropertyChanged "FillColor"
txtParent.BackColor = FillColor
End Property
