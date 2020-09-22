VERSION 5.00
Begin VB.UserControl XMultiButton 
   ClientHeight    =   555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1620
   ScaleHeight     =   555
   ScaleWidth      =   1620
   ToolboxBitmap   =   "XClick.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1200
      Top             =   0
   End
   Begin VB.Label lblMulti 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XMultiButton"
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Shape cmdMulti 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "XMultiButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'API Declarations:
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
'Default Property Values:
Const m_def_ClickColor = &HFF8080
Const m_def_FillColor = vbBlack
Const m_def_MouseEnterColor = &H800000
'Property Variables:
Dim m_ClickColor As OLE_COLOR
Dim m_FillColor As OLE_COLOR
Dim m_MouseEnterColor As OLE_COLOR
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."

Private Function UnderMouse() As Boolean
    Dim ptMouse As POINTAPI
    GetCursorPos ptMouse
    If WindowFromPoint(ptMouse.X, ptMouse.Y) = UserControl.hWnd Then
       UnderMouse = True
    Else
       UnderMouse = False
    End If
End Function

Private Function Runtime()
If Ambient.UserMode Then
       Timer1.Enabled = True
    Else: Timer1.Enabled = False
End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblMulti.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblMulti.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMulti,cmdMulti,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = cmdMulti.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    cmdMulti.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMulti,cmdMulti,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    cmdMulti.Refresh
End Sub

Private Sub lblMulti_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsDown
    Timer1.Enabled = False
End Sub
'
'Private Sub lblMulti_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    RaiseEvent MouseMove(Button, Shift, X, Y)
'End Sub

Private Sub lblMulti_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MsUp
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If UnderMouse = True Then
   cmdMulti.BackColor = MouseEnterColor
   Else
   cmdMulti.BackColor = FillColor
End If
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblMulti.Caption
    UserControl_Resize
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblMulti.Caption() = New_Caption
    PropertyChanged "Caption"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lblMulti.FontBold
    UserControl_Resize
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblMulti.FontBold() = New_FontBold
    PropertyChanged "FontBold"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lblMulti.FontItalic
    UserControl_Resize
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblMulti.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lblMulti.FontName
    UserControl_Resize
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblMulti.FontName() = New_FontName
    PropertyChanged "FontName"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lblMulti.FontSize
    UserControl_Resize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblMulti.FontSize() = New_FontSize
    PropertyChanged "FontSize"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = lblMulti.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblMulti.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = lblMulti.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblMulti.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MouseEnterColor() As OLE_COLOR
    MouseEnterColor = m_MouseEnterColor
End Property

Public Property Let MouseEnterColor(ByVal New_MouseEnterColor As OLE_COLOR)
    m_MouseEnterColor = New_MouseEnterColor
    PropertyChanged "MouseEnterColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    lblMulti.Caption = Extender.Name
    m_MouseEnterColor = m_def_MouseEnterColor
    m_FillColor = m_def_FillColor
    m_ClickColor = m_def_ClickColor
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    Timer1.Enabled = False
    MsDown
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    Timer1.Enabled = True
    MsUp
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblMulti.Font = PropBag.ReadProperty("Font", Ambient.Font)
    cmdMulti.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    lblMulti.Caption = PropBag.ReadProperty("Caption", "XMultiButton")
    lblMulti.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblMulti.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblMulti.FontName = PropBag.ReadProperty("FontName", "")
    lblMulti.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblMulti.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    lblMulti.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblMulti.ForeColor = PropBag.ReadProperty("ForeColor", vbYellow)
    m_MouseEnterColor = PropBag.ReadProperty("MouseEnterColor", m_def_MouseEnterColor)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_ClickColor = PropBag.ReadProperty("ClickColor", m_def_ClickColor)
    cmdMulti.BorderColor = PropBag.ReadProperty("BorderColor", &HFF0000)
    cmdMulti.BorderWidth = PropBag.ReadProperty("BorderWidth", 1)
    cmdMulti.BackColor = FillColor
    UserControl_Resize
End Sub
Private Sub UserControl_Resize()
    cmdMulti.Left = 0
    cmdMulti.Top = 0
    lblMulti.Left = 0
    lblMulti.Top = 0
    cmdMulti.Height = UserControl.Height
    cmdMulti.Width = UserControl.Width
    lblMulti.Move (ScaleWidth - lblMulti.Width) / 2, (ScaleHeight - lblMulti.Height) / 2, lblMulti.Width
End Sub
'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblMulti.Font, Ambient.Font)
    Call PropBag.WriteProperty("BorderStyle", cmdMulti.BorderStyle, 1)
    Call PropBag.WriteProperty("Caption", lblMulti.Caption, "XMultiButton")
    Call PropBag.WriteProperty("FontBold", lblMulti.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblMulti.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblMulti.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblMulti.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblMulti.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", lblMulti.FontUnderline, 0)
    Call PropBag.WriteProperty("MouseEnterColor", m_MouseEnterColor, m_def_MouseEnterColor)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("ClickColor", m_ClickColor, m_def_ClickColor)
    Call PropBag.WriteProperty("ForeColor", lblMulti.ForeColor, vbYellow)
    Call PropBag.WriteProperty("BorderColor", cmdMulti.BorderColor, &HFF0000)
    Call PropBag.WriteProperty("BorderWidth", cmdMulti.BorderWidth, 1)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
    cmdMulti.BackColor = FillColor
End Property

Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_FillColor = New_FillColor
    PropertyChanged "FillColor"
    cmdMulti.BackColor = FillColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ClickColor() As OLE_COLOR
    ClickColor = m_ClickColor
End Property

Public Property Let ClickColor(ByVal New_ClickColor As OLE_COLOR)
    m_ClickColor = New_ClickColor
    PropertyChanged "ClickColor"
End Property

Private Sub MsDown()
    cmdMulti.BackColor = ClickColor
End Sub

Private Sub MsUp()
    RaiseEvent Click
    cmdMulti.BackColor = FillColor
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblMulti,lblMulti,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblMulti.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblMulti.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMulti,cmdMulti,-1,BorderColor
Public Property Get BorderColor() As OLE_COLOR
Attribute BorderColor.VB_Description = "Returns/sets the color of an object's border."
    BorderColor = cmdMulti.BorderColor
End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)
    cmdMulti.BorderColor() = New_BorderColor
    PropertyChanged "BorderColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=cmdMulti,cmdMulti,-1,BorderWidth
Public Property Get BorderWidth() As Integer
Attribute BorderWidth.VB_Description = "Returns or sets the width of a control's border."
    BorderWidth = cmdMulti.BorderWidth
End Property

Public Property Let BorderWidth(ByVal New_BorderWidth As Integer)
    cmdMulti.BorderWidth() = New_BorderWidth
    PropertyChanged "BorderWidth"
End Property

