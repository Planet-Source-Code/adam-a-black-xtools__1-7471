VERSION 5.00
Begin VB.UserControl XTransButton 
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1485
   ScaleHeight     =   615
   ScaleWidth      =   1485
   ToolboxBitmap   =   "XTransButton.ctx":0000
   Begin VB.Label lblCmd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XButton"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   570
   End
   Begin VB.Shape XButton 
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "XTransButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim Color As Long
Dim m_FontColor As OLE_COLOR
Dim m_ClickColor As OLE_COLOR
Const m_def_ClickColor = &HFFE3B5
Const m_def_FontColor = &H0&
Event Click()
Event DblClick()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Sub MsDown()
XButton.FillColor = ClickColor
End Sub
Private Sub MsUp()
RaiseEvent Click
XButton.FillColor = Color
End Sub
Private Sub lblCmd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MsDown
End Sub
Private Sub lblCmd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call MsUp
End Sub
Private Sub UserControl_InitProperties()
lblCmd.Caption = Extender.Name
m_ClickColor = m_def_ClickColor
m_FontColor = m_def_FontColor
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
Call MsDown
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
Call MsUp
End Sub
Private Sub UserControl_Resize()
Color = XButton.FillColor
XButton.Left = 0
XButton.Top = 0
lblCmd.Left = 0
lblCmd.Top = 0
XButton.Height = UserControl.Height
XButton.Width = UserControl.Width
lblCmd.Move (ScaleWidth - lblCmd.Width) / 2, (ScaleHeight - lblCmd.Height) / 2, lblCmd.Width
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
lblCmd.Caption = PropBag.ReadProperty("Caption", "XButton")
XButton.FillColor = PropBag.ReadProperty("FillColor", &HC0C0C0)
XButton.Shape = PropBag.ReadProperty("Shape", 4)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
m_ClickColor = PropBag.ReadProperty("ClickColor", m_def_ClickColor)
m_FontColor = PropBag.ReadProperty("FontColor", m_def_FontColor)
Set lblCmd.Font = PropBag.ReadProperty("Font", Ambient.Font)
UserControl_Resize
New_Color
End Sub
Public Property Get Caption() As Variant
Caption = lblCmd.Caption
End Property
Public Property Let Caption(ByVal vNewValue As Variant)
lblCmd.Caption = vNewValue
UserControl_Resize
End Property
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", lblCmd.Caption
Call PropBag.WriteProperty("FillColor", XButton.FillColor, &HC0C0C0)
Call PropBag.WriteProperty("Shape", XButton.Shape, 4)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
Call PropBag.WriteProperty("ClickColor", m_ClickColor, m_def_ClickColor)
Call PropBag.WriteProperty("Font", lblCmd.Font, Ambient.Font)
Call PropBag.WriteProperty("FontColor", m_FontColor, m_def_FontColor)
End Sub
Public Property Get FillColor() As OLE_COLOR
FillColor = XButton.FillColor
Color = XButton.FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
XButton.FillColor() = New_FillColor
PropertyChanged "FillColor"
Color = XButton.FillColor
End Property
Public Property Get Shape() As Integer
Shape = XButton.Shape
End Property
Public Property Let Shape(ByVal New_Shape As Integer)
XButton.Shape() = New_Shape
PropertyChanged "Shape"
End Property
Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub
Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
UserControl.Enabled() = New_Enabled
PropertyChanged "Enabled"
End Property
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub
Public Property Get MousePointer() As Integer
MousePointer = UserControl.MousePointer
End Property
Public Property Let MousePointer(ByVal New_MousePointer As Integer)
UserControl.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property
Private Sub UserControl_Click()
RaiseEvent Click
End Sub
Public Property Get ClickColor() As OLE_COLOR
ClickColor = m_ClickColor
End Property
Public Property Let ClickColor(ByVal New_ClickColor As OLE_COLOR)
m_ClickColor = New_ClickColor
PropertyChanged "ClickColor"
End Property
Public Property Get Font() As Font
Set Font = lblCmd.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
Set lblCmd.Font = New_Font
PropertyChanged "Font"
UserControl_Resize
End Property
Public Property Get FontColor() As OLE_COLOR
FontColor = m_FontColor
End Property
Public Property Let FontColor(ByVal New_FontColor As OLE_COLOR)
m_FontColor = New_FontColor
PropertyChanged "FontColor"
New_Color
End Property
Private Sub New_Color()
lblCmd.ForeColor = FontColor
End Sub

