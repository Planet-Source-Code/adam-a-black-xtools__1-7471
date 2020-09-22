VERSION 5.00
Begin VB.UserControl XCheck 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ScaleHeight     =   300
   ScaleWidth      =   2535
   ToolboxBitmap   =   "XCheck.ctx":0000
   Begin VB.CheckBox chkParent 
      Caption         =   "XCheck"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "XCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function SetCapture Lib _
"user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" _
() As Long
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
ForeColor = chkParent.ForeColor
End Property
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
chkParent.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"
End Property
Public Property Get Enabled() As Boolean
Enabled = chkParent.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
chkParent.Enabled() = New_Enabled
PropertyChanged "Enabled"
End Property
Public Property Get Font() As Font
Set Font = chkParent.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
Set chkParent.Font = New_Font
PropertyChanged "Font"
End Property
Public Sub Refresh()
chkParent.Refresh
End Sub
Private Sub chkParent_Click()
RaiseEvent Click
End Sub
Private Sub chkParent_DblClick()
RaiseEvent DblClick
End Sub
Private Sub chkParent_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub chkParent_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub chkParent_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub chkParent_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub
Private Sub chkParent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
If FocusState = False Then
With chkParent
If (X < 0) Or (Y < 0) Or (X > .Width) Or (Y > .Height) Then
ReleaseCapture
chkParent.BackColor = FillColor
Else
SetCapture .hWnd
chkParent.BackColor = FocusColor
End If
End With
End If
End Sub
Private Sub chkParent_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub
Public Property Get Alignment() As Integer
Alignment = chkParent.Alignment
End Property
Public Property Let Alignment(ByVal New_Alignment As Integer)
chkParent.Alignment() = New_Alignment
PropertyChanged "Alignment"
End Property
Private Sub chkParent_Change()
RaiseEvent Change
End Sub
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
chkParent.BackColor = FocusColor
End If
End Sub
Private Sub UserControl_ExitFocus()
FocusState = False
If FocusState = False Then
chkParent.BackColor = FillColor
End If
End Sub
Private Sub UserControl_InitProperties()
chkParent.Caption = Extender.Name
m_FocusColor = m_def_FocusColor
m_FillColor = m_def_FillColor
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
chkParent.Caption = PropBag.ReadProperty("Caption", "")
chkParent.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
chkParent.Enabled = PropBag.ReadProperty("Enabled", True)
Set chkParent.Font = PropBag.ReadProperty("Font", Ambient.Font)
chkParent.Alignment = PropBag.ReadProperty("Alignment", 0)
m_FocusColor = PropBag.ReadProperty("FocusColor", m_def_FocusColor)
m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
chkParent.BackColor = FillColor
chkParent.Value = PropBag.ReadProperty("Value", 0)
End Sub
Private Sub UserControl_Resize()
chkParent.Height = UserControl.Height
chkParent.Width = UserControl.Width
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Caption", chkParent.Caption, "")
Call PropBag.WriteProperty("ForeColor", chkParent.ForeColor, &H80000008)
Call PropBag.WriteProperty("Enabled", chkParent.Enabled, True)
Call PropBag.WriteProperty("Font", chkParent.Font, Ambient.Font)
Call PropBag.WriteProperty("Alignment", chkParent.Alignment, 0)
Call PropBag.WriteProperty("FocusColor", m_FocusColor, m_def_FocusColor)
Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
Call PropBag.WriteProperty("Value", chkParent.Value, 0)
End Sub
Public Property Get FillColor() As OLE_COLOR
FillColor = m_FillColor
chkParent.BackColor = FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
m_FillColor = New_FillColor
PropertyChanged "FillColor"
chkParent.BackColor = FillColor
End Property
Public Property Get Value() As Integer
Attribute Value.VB_Description = "Returns/sets the value of an object."
    Value = chkParent.Value
End Property
Public Property Let Value(ByVal New_Value As Integer)
    chkParent.Value() = New_Value
    PropertyChanged "Value"
End Property
Public Property Get Caption() As String
Caption = chkParent.Caption
End Property
Public Property Let Caption(ByVal New_Caption As String)
chkParent.Caption() = New_Caption
PropertyChanged "Caption"
End Property

