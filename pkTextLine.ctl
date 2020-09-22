VERSION 5.00
Begin VB.UserControl XMText 
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1770
   LockControls    =   -1  'True
   ScaleHeight     =   390
   ScaleWidth      =   1770
   ToolboxBitmap   =   "pkTextLine.ctx":0000
   Begin VB.TextBox txtParent 
      BackColor       =   &H00FFF5E8&
      BorderStyle     =   0  'None
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1065
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   1080
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000014&
      X1              =   15
      X2              =   1200
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000014&
      X1              =   1200
      X2              =   1200
      Y1              =   495
      Y2              =   0
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   1200
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "XMText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'API Declarations
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
'Default Property Values:
Dim PixelX
Dim PixelY
Dim m_PasswordChr As Variant
Dim m_Font As Font
Dim KG As Integer
Dim FocusState As Boolean
Dim m_ChangeColor As Boolean
Dim m_FillColor As OLE_COLOR
Dim m_FocusColor As OLE_COLOR
'Property Values
Const m_def_FillColor = &HFFE3B5
Const m_def_ChangeColor = True
Const m_def_PasswordChar = 0
Const m_def_FocusColor = vbWhite
'Event Declarations():
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Change()
Public Enum PasswordChar
No_Char = 0
Stars = 1
End Enum

Private Sub txtParent_GotFocus()
FocusState = True
If FocusState = True And ChangeColor = True Then
txtParent.BackColor = FocusColor
End If
End Sub

Private Sub txtParent_LostFocus()
FocusState = False
If FocusState = False And ChangeColor = True Then
txtParent.BackColor = FillColor
End If
End Sub

Private Sub UserControl_Initialize()
 PixelX = Screen.TwipsPerPixelX
 PixelY = Screen.TwipsPerPixelY
 Line5.BorderColor = txtParent.BackColor
 Line3.Y2 = -PixelY
 UserControl_Resize
End Sub

Private Sub Font_Change()
 txtParent.Height = txtParent.Height + PixelX
 UserControl.Height = txtParent.Height + 4 * PixelY
 UserControl.Width = txtParent.Width + 4 * PixelX
 UserControl_Resize
End Sub

Private Sub UserControl_Resize()
 If UserControl.Width < 10 * PixelY Then
  UserControl.Width = 10 * PixelY
 End If
 

 If UserControl.Height <= 17 * PixelY Then
  UserControl.Height = 17 * PixelY
  KG = 0
  Line5.Visible = False
 ElseIf UserControl.Height < 19 * PixelY Then
  UserControl.Height = 19 * PixelY
  KG = 1
  Line5.Visible = True
 Else
  KG = 1
  Line5.Visible = True
 End If
 
 txtParent.Height = UserControl.Height - PixelX * (4 + KG)
 txtParent.Width = UserControl.Width - PixelY * 4
 txtParent.Top = PixelX * (2 + KG)
 txtParent.Left = PixelY
 txtParent.Left = PixelX * 2
 
 Line1.Y2 = UserControl.Height
 Line2.X2 = UserControl.Width - PixelX
 Line3.X1 = UserControl.Width - PixelX
 Line3.X2 = UserControl.Width - PixelX
 Line3.Y1 = UserControl.Height
 Line4.X2 = UserControl.Width - PixelX
 Line4.Y1 = UserControl.Height - PixelY
 Line4.Y2 = UserControl.Height - PixelY
 Line5.Y1 = txtParent.Top - PixelY
 Line5.Y2 = txtParent.Top - PixelY
 Line5.X1 = txtParent.Left
 Line5.X2 = txtParent.Left + txtParent.Width
 Line5.BorderColor = txtParent.BackColor
End Sub

Public Property Get ForeColor() As OLE_COLOR
 ForeColor = txtParent.ForeColor
End Property

Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
 txtParent.ForeColor = vNewValue
 PropertyChanged "ForeColor"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", txtParent.ForeColor, &H80000008)
    Call PropBag.WriteProperty("Text", txtParent.Text, "")
    Call PropBag.WriteProperty("Enabled", txtParent.Enabled, True)
    Call PropBag.WriteProperty("Font", Font, Ambient.Font)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Enabled", txtParent.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("Text", txtParent.Text, "")
    Call PropBag.WriteProperty("MaxLength", txtParent.MaxLength, 0)
    Call PropBag.WriteProperty("ToolTipText", txtParent.ToolTipText, "")
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("ColorHighLight", Line3.BorderColor, -2147483628)
    Call PropBag.WriteProperty("ColorDarkShadow", Line1.BorderColor, -2147483632)
    Call PropBag.WriteProperty("FrameColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("FocusColor", m_FocusColor, m_def_FocusColor)
    Call PropBag.WriteProperty("FillColor", m_FillColor, m_def_FillColor)
    Call PropBag.WriteProperty("ChangeColor", m_ChangeColor, m_def_ChangeColor)
    Call PropBag.WriteProperty("PasswordChr", m_PasswordChr, m_def_PasswordChar)
    Call PropBag.WriteProperty("SelLength", txtParent.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", txtParent.SelStart, 0)
    Call PropBag.WriteProperty("SelText", txtParent.SelText, "")
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    txtParent.SelLength = PropBag.ReadProperty("SelLength", 0)
    txtParent.SelStart = PropBag.ReadProperty("SelStart", 0)
    txtParent.SelText = PropBag.ReadProperty("SelText", "")
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    Set txtParent.Font = PropBag.ReadProperty("Font", Ambient.Font)
    txtParent.Text = PropBag.ReadProperty("Text", "")
    Line5.BorderColor = PropBag.ReadProperty("BackColor", &H80000005)
    txtParent.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    txtParent.Enabled = PropBag.ReadProperty("Enabled", True)
    txtParent.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtParent.Text = PropBag.ReadProperty("Text", "")
    txtParent.MaxLength = PropBag.ReadProperty("MaxLength", 0)
    txtParent.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Line3.BorderColor = PropBag.ReadProperty("ColorHighLight", -2147483628)
    UserControl.BackColor = PropBag.ReadProperty("ColorLightShadow", &H8000000F)
    Line1.BorderColor = PropBag.ReadProperty("ColorDarkShadow", -2147483632)
    UserControl.BackColor = PropBag.ReadProperty("FrameColor", &H8000000F)
    m_FocusColor = PropBag.ReadProperty("FocusColor", m_def_FocusColor)
    m_FillColor = PropBag.ReadProperty("FillColor", m_def_FillColor)
    m_ChangeColor = PropBag.ReadProperty("ChangeColor", m_def_ChangeColor)
    m_PasswordChr = PropBag.ReadProperty("PasswordChr", m_def_PasswordChar)
    PswdChar
    txtParent.BackColor = FillColor
End Sub

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=txtParent,txtParent,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = txtParent.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    txtParent.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hDC() As Long
    hDC = UserControl.hDC
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'MappingInfo=txtParent,txtParent,-1,Text
Public Property Get Text() As String
    Text = txtParent.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    txtParent.Text() = New_Text
    PropertyChanged "Text"
End Property

Private Sub txtParent_Change()
    RaiseEvent Change
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'MappingInfo=txtParent,txtParent,-1,MaxLength
Public Property Get MaxLength() As Long
    MaxLength = txtParent.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
    txtParent.MaxLength() = New_MaxLength
    PropertyChanged "MaxLength"
End Property

Private Sub txtParent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
    If FocusState = False And ChangeColor = True Then
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

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'MappingInfo=txtParent,txtParent,-1,ToolTipText
Public Property Get ToolTipText() As String
    ToolTipText = txtParent.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    txtParent.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Public Property Get Font() As Font
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Private Sub UserControl_InitProperties()
    Set m_Font = Ambient.Font
    txtParent.Text = Extender.Name
    m_ChangeColor = m_def_ChangeColor
    m_FocusColor = m_def_FocusColor
    m_FillColor = m_def_FillColor
End Sub

Public Property Get ChangeColor() As Boolean
    ChangeColor = m_ChangeColor
End Property

Public Property Let ChangeColor(ByVal New_ChangeColor As Boolean)
    m_ChangeColor = New_ChangeColor
    PropertyChanged "ChangeColor"
End Property

'ACHTUNG! DIE FOLGENDEN KOMMENTIERTEN ZEILEN NICHT ENTFERNEN ODER VERÄNDERN!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get FrameColor() As OLE_COLOR
    FrameColor = UserControl.BackColor
End Property

Public Property Let FrameColor(ByVal New_FrameColor As OLE_COLOR)
    UserControl.BackColor() = New_FrameColor
    PropertyChanged "FrameColor"
End Property

Public Property Get FillColor() As OLE_COLOR
    FillColor = m_FillColor
    txtParent.BackColor = FillColor
End Property
Public Property Let FillColor(ByVal New_FillColor As OLE_COLOR)
    m_FillColor = New_FillColor
    PropertyChanged "FillColor"
    txtParent.BackColor = FillColor
End Property

Public Property Get FocusColor() As OLE_COLOR
    FocusColor = m_FocusColor
End Property
Public Property Let FocusColor(ByVal New_FocusColor As OLE_COLOR)
    m_FocusColor = New_FocusColor
    PropertyChanged "FocusColor"
End Property

Public Property Get PasswordChr() As PasswordChar
    PasswordChr = m_PasswordChr
    PswdChar
End Property

Public Property Let PasswordChr(ByVal New_PasswordChr As PasswordChar)
    m_PasswordChr = New_PasswordChr
    PswdChar
    PropertyChanged "PasswordChr"
End Property

Private Sub PswdChar()
Select Case m_PasswordChr
    Case 0
    txtParent.PasswordChar = ""
    Case 1
    txtParent.PasswordChar = "*"
    Case Else
    Exit Sub
End Select
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtParent,txtParent,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = txtParent.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    txtParent.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtParent,txtParent,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = txtParent.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    txtParent.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtParent,txtParent,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = txtParent.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    txtParent.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

