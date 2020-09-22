VERSION 5.00
Begin VB.UserControl XTickBox 
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   Picture         =   "XTickBox.ctx":0000
   ScaleHeight     =   240
   ScaleWidth      =   1545
   ToolboxBitmap   =   "XTickBox.ctx":0046
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   360
   End
   Begin VB.PictureBox imgCheck 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "XTickBox.ctx":0358
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
   Begin VB.Image imgCross 
      Appearance      =   0  'Flat
      Height          =   150
      Left            =   720
      Picture         =   "XTickBox.ctx":04B2
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image imgTick 
      Height          =   150
      Left            =   480
      Picture         =   "XTickBox.ctx":060C
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblTick 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "XTickBox"
      Height          =   195
      Left            =   250
      TabIndex        =   1
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "XTickBox"
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
Const m_def_CheckStyle = 0
Const m_def_Value = 0
Const m_def_MouseInColor = vbYellow
Const m_def_MouseOutColor = &HFFF5E8
'Property Variables:
Dim Img As String
Dim m_CheckStyle As Variant
Dim m_Value As Variant
Dim m_MouseInColor As OLE_COLOR
Dim m_MouseOutColor As OLE_COLOR
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Change() 'MappingInfo=lblTick,lblTick,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,OLEDragDrop
Attribute OLEDragDrop.VB_Description = "Occurs when data is dropped onto the control via an OLE drag/drop operation, and OLEDropMode is set to manual."
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer) 'MappingInfo=UserControl,UserControl,-1,OLEDragOver
Attribute OLEDragOver.VB_Description = "Occurs when the mouse is moved over the control during an OLE drag/drop operation, if its OLEDropMode property is set to manual."
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long) 'MappingInfo=UserControl,UserControl,-1,OLEStartDrag
Attribute OLEStartDrag.VB_Description = "Occurs when an OLE drag/drop operation is initiated either manually or automatically."
Event OLESetData(Data As DataObject, DataFormat As Integer) 'MappingInfo=UserControl,UserControl,-1,OLESetData
Attribute OLESetData.VB_Description = "Occurs at the OLE drag/drop source control when the drop target requests data that was not provided to the DataObject during the OLEDragStart event."
'Enumerated Constants:
Public Enum Value
Unchecked = 0
Checked = 1
End Enum
Public Enum CheckStyle
Tick = 0
Cross = 1
End Enum

Private Sub imgCheck_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub imgCheck_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
        NewValue
    End Select
End Sub

Private Sub imgCheck_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub imgCheck_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub imgCheck_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then
        NewValue
    End If
    RaiseEvent Click
End Sub

Private Sub imgCheck_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub imgCheck_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub imgCheck_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub imgCheck_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub lblTick_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub lblTick_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lblTick_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lblTick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then
        NewValue
    End If
    RaiseEvent Click
End Sub

Private Sub lblTick_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub lblTick_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub lblTick_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub lblTick_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Timer1_Timer()
If UnderMouse = True Then
   imgCheck.BackColor = MouseInColor
   Else
   imgCheck.BackColor = MouseOutColor
End If
End Sub

Private Sub UserControl_Initialize()
    UserControl_Resize
    Value = Unchecked
    UserControl.Picture = LoadPicture()
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
        NewValue
    End Select
End Sub

Private Sub UserControl_Resize()
    lblTick.Top = 0
    lblTick.Left = 0
    imgCheck.Top = 0
    imgCheck.Left = 0
    lblTick.Top = (UserControl.Height - lblTick.Height) / 2
    imgCheck.Top = (UserControl.Height - imgCheck.Height) / 2
    lblTick.Left = (imgCheck.Width) + 80
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = lblTick.BackColor
    UserControl.BackColor = lblTick.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblTick.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    UserControl.BackColor = lblTick.BackColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = lblTick.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblTick.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

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
'MappingInfo=lblTick,lblTick,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblTick.Font
    UserControl_Resize
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblTick.Font = New_Font
    PropertyChanged "Font"
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    UserControl.Refresh
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button = 1 Then
        NewValue
    End If
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Caption = lblTick.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    lblTick.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

Private Sub lblTick_Change()
    RaiseEvent Change
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = lblTick.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    lblTick.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = lblTick.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    lblTick.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = lblTick.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    lblTick.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = lblTick.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    lblTick.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = lblTick.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    lblTick.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTick,lblTick,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = lblTick.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    lblTick.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
Attribute hWnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDrag
Public Sub OLEDrag()
Attribute OLEDrag.VB_Description = "Starts an OLE drag/drop event with the given control as the source."
    UserControl.OLEDrag
End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,OLEDropMode
Public Property Get OLEDropMode() As Integer
Attribute OLEDropMode.VB_Description = "Returns/Sets whether this object can act as an OLE drop target."
    OLEDropMode = UserControl.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
    UserControl.OLEDropMode() = New_OLEDropMode
    PropertyChanged "OLEDropMode"
End Property

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MouseInColor() As OLE_COLOR
Attribute MouseInColor.VB_Description = "Returns/sets a graphic to be displayed in a control."
    MouseInColor = m_MouseInColor
End Property

Public Property Let MouseInColor(ByVal New_MouseInColor As OLE_COLOR)
    m_MouseInColor = New_MouseInColor
    PropertyChanged "MouseInColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get MouseOutColor() As OLE_COLOR
    MouseOutColor = m_MouseOutColor
    imgCheck.BackColor = MouseOutColor
End Property

Public Property Let MouseOutColor(ByVal New_MouseOutColor As OLE_COLOR)
    m_MouseOutColor = New_MouseOutColor
    imgCheck.BackColor = MouseOutColor
    PropertyChanged "MouseOutColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    lblTick.Caption = Extender.Name
    m_MouseInColor = m_def_MouseInColor
    m_MouseOutColor = m_def_MouseOutColor
    m_Value = m_def_Value
    m_CheckStyle = m_def_CheckStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblTick.BackColor = PropBag.ReadProperty("BackColor", &HC0C0C0)
    lblTick.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblTick.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    lblTick.Caption = PropBag.ReadProperty("Caption", "XTickBox")
    lblTick.FontBold = PropBag.ReadProperty("FontBold", 0)
    lblTick.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    lblTick.FontName = PropBag.ReadProperty("FontName", "")
    lblTick.FontSize = PropBag.ReadProperty("FontSize", 0)
    lblTick.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    lblTick.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    UserControl.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
    m_CheckStyle = PropBag.ReadProperty("CheckStyle", m_def_CheckStyle)
    m_MouseInColor = PropBag.ReadProperty("MouseInColor", m_def_MouseInColor)
    m_MouseOutColor = PropBag.ReadProperty("MouseOutColor", m_def_MouseOutColor)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    UserControl.BackColor = lblTick.BackColor
    imgCheck.BackColor = MouseOutColor
    UserControl_Resize
    GetImage
    ValCheck
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", lblTick.BackColor, &HC0C0C0)
    Call PropBag.WriteProperty("ForeColor", lblTick.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", lblTick.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("Caption", lblTick.Caption, "XTickBox")
    Call PropBag.WriteProperty("FontBold", lblTick.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", lblTick.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", lblTick.FontName, "")
    Call PropBag.WriteProperty("FontSize", lblTick.FontSize, 0)
    Call PropBag.WriteProperty("FontUnderline", lblTick.FontUnderline, 0)
    Call PropBag.WriteProperty("FontStrikethru", lblTick.FontStrikethru, 0)
    Call PropBag.WriteProperty("OLEDropMode", UserControl.OLEDropMode, 0)
    Call PropBag.WriteProperty("MouseInColor", m_MouseInColor, m_def_MouseInColor)
    Call PropBag.WriteProperty("MouseOutColor", m_MouseOutColor, m_def_MouseOutColor)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("CheckStyle", m_CheckStyle, m_def_CheckStyle)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get Value() As Value
    Value = m_Value
    ValCheck
End Property

Public Property Let Value(ByVal New_Value As Value)
    m_Value = New_Value
    ValCheck
    RaiseEvent Change
    PropertyChanged "Value"
End Property

Private Function Runtime()
If Ambient.UserMode = True Then
    Timer1.Enabled = True
    Else
    Timer1.Enabled = False
End If
End Function

Private Sub NewValue()
Select Case m_Value
    Case 0
    If Img = "Tick" Then
    imgCheck.Picture = imgTick.Picture
    Else
    imgCheck.Picture = imgCross.Picture
    End If
    m_Value = 1
    Case 1
    imgCheck.Picture = LoadPicture()
    m_Value = 0
End Select
RaiseEvent Change
End Sub

Private Sub ValCheck()
Select Case m_Value
    Case 0
    imgCheck.Picture = LoadPicture()
    m_Value = 0
    Case 1
    If Img = "Tick" Then
    imgCheck.Picture = imgTick.Picture
    m_Value = 1
    Else
    imgCheck.Picture = imgCross.Picture
    m_Value = 1
    End If
End Select
End Sub

Private Function UnderMouse() As Boolean
    Dim ptMouse As POINTAPI
    GetCursorPos ptMouse
    If WindowFromPoint(ptMouse.X, ptMouse.Y) = UserControl.hWnd Or WindowFromPoint(ptMouse.X, ptMouse.Y) = imgCheck.hWnd Then
       UnderMouse = True
    Else
       UnderMouse = False
    End If
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CheckStyle() As CheckStyle
    CheckStyle = m_CheckStyle
    GetImage
End Property

Public Property Let CheckStyle(ByVal New_CheckStyle As CheckStyle)
    m_CheckStyle = New_CheckStyle
    GetImage
    PropertyChanged "CheckStyle"
End Property

Private Sub GetImage()
Select Case m_CheckStyle
    Case 0
    Img = "Tick"
    Case 1
    Img = "Cross"
End Select
End Sub

