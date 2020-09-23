VERSION 5.00
Begin VB.UserControl Position 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "Position.ctx":0000
   Begin VB.Timer DragTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4200
      Top             =   0
   End
   Begin VB.Image Cursor 
      Appearance      =   0  'Flat
      Height          =   180
      Left            =   180
      Top             =   60
      Width           =   135
   End
End
Attribute VB_Name = "Position"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Windows API functions
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoints As Any, ByVal nCount As Long) As Long
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
'Private types
Private Type Point
  X As Long
  Y As Long
End Type
'Private properties
Private ControlOK As Boolean
Private TickPos(1 To 255) As Long
Private DragCursor As Boolean
'Default Property Values:
Const ForeColorDefault = 0
Const PositionsDefault = 10
Const ValueDefault = 1
'Property Variables:
Private mForeColor As OLE_COLOR
Private mPositions As Byte
Private mValue As Long    'default property
'Public Events
Public Event Change()


'----------------------------------------------- PRIVATE
Private Sub DrawControl()
  Dim Space As Long
  Dim Excesso As Long
  Dim i As Long
  Dim p1X As Long
  Dim p1Y As Long
  Dim p2X As Long
  Dim p2Y As Long
  Dim PolP(0 To 2) As Point
  
  UserControl.ForeColor = mForeColor
  'controlo preto
  If ControlOK = False Then
    Line (0, 0)-(UserControl.ScaleWidth, UserControl.ScaleHeight), mForeColor, BF
    Exit Sub
  End If
  'desenhar ticks
  UserControl.Cls
  If Positions = 1 Then Space = 0 Else Space = (UserControl.ScaleWidth - 20) / (Positions - 1)
  Excesso = UserControl.ScaleWidth - 20 - Space * (Positions - 1)
  For i = 1 To Positions
    TickPos(i) = 10 + Space * (i - 1) + Excesso / 2
    Line (TickPos(i), UserControl.ScaleHeight - 3)-Step(0, 3)
  Next i
  'desenhar buraco do slider
  p1X = TickPos(1) - 3
  p1Y = UserControl.ScaleHeight / 2 - 1
  p2X = TickPos(Positions) + 3
  UserControl.ForeColor = vb3DHighlight
  Line (p1X, p1Y + 2)-(p2X, p1Y + 2)
  Line (p2X, p1Y + 2)-(p2X, p1Y - 3)
  UserControl.ForeColor = vb3DFace
  Line (p1X + 1, p1Y + 1)-(p2X - 1, p1Y + 1)
  Line (p2X - 1, p1Y + 1)-(p2X - 1, p1Y - 2)
  UserControl.ForeColor = vb3DShadow
  Line (p1X, p1Y + 1)-(p1X, p1Y - 2)
  Line (p1X, p1Y - 2)-(p2X, p1Y - 2)
  UserControl.ForeColor = 0
  Line (p1X + 1, p1Y - 1)-(p2X - 2, p1Y), 0, BF
  'desenhar slider
  p1X = TickPos(mValue)
  p1Y = UserControl.ScaleHeight / 2
  UserControl.ForeColor = vb3DFace
  PolP(0).X = p1X - 7:  PolP(0).Y = p1Y - 7
  PolP(1).X = p1X + 7:  PolP(1).Y = p1Y - 7
  PolP(2).X = p1X:      PolP(2).Y = p1Y + 7
  Call Polygon(UserControl.hDC, PolP(0), 3)
  UserControl.ForeColor = vb3DHighlight
  Line (p1X - 7, p1Y - 7)-(p1X + 7, p1Y - 7)
  Line (p1X - 7, p1Y - 7)-(p1X, p1Y + 8)
  UserControl.ForeColor = vb3DShadow
  Line (p1X + 7, p1Y - 7)-(p1X, p1Y + 7)
  UserControl.ForeColor = vb3DDKShadow
  Line (p1X + 8, p1Y - 7)-(p1X + 1, p1Y + 8)
  'posicionar cursor (imagem) no lugar correcto
  Cursor.top = p1Y - 6
  Cursor.Left = p1X - 4
  'indicação do valor em que está
  Me.ToolTipText = mValue
End Sub

'----------------------------------------------- PROPERTIES
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As MSComctlLib.BorderStyleConstants
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As MSComctlLib.BorderStyleConstants)
  UserControl.BorderStyle() = New_BorderStyle
  PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get Value() As Long
Attribute Value.VB_UserMemId = 0
  Value = mValue
End Property

Public Property Let Value(ByVal New_Value As Long)
  If (New_Value >= 1) And (New_Value <= Positions) Then mValue = New_Value
  PropertyChanged "Value"
  Call DrawControl
  RaiseEvent Change
End Property

Private Sub Cursor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If ControlOK = False Then Exit Sub
  DragCursor = True
  DragTimer.Enabled = True
End Sub

Private Sub Cursor_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  DragCursor = False
  DragTimer.Enabled = False
  Call DragTimer_Timer
End Sub

Private Sub DragTimer_Timer()
  Dim Mpos As Point
  Dim Wpos As Point
  Dim PosX As Long
  Dim i As Long
  Dim Min As Long
  Dim Valor As Byte
  
  If DragCursor = False Then DragTimer.Enabled = False
  'determina posição
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Call ClientToScreen(UserControl.hWnd, Wpos)
  PosX = Mpos.X - Wpos.X
  'determina valor pela posição
  Min = 999999
  For i = 1 To Positions
    If Abs(TickPos(i) - PosX) <= Min Then
      Min = Abs(TickPos(i) - PosX)
      Valor = i
    End If
  Next i
  If mValue <> Valor Then
    mValue = Valor
    Call DrawControl
    RaiseEvent Change
  End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
  ControlOK = False
  mValue = ValueDefault
  mPositions = PositionsDefault
  mForeColor = ForeColorDefault
  DragCursor = False
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

  UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
  mValue = PropBag.ReadProperty("Value", ValueDefault)
  mPositions = PropBag.ReadProperty("Positions", PositionsDefault)
  mForeColor = PropBag.ReadProperty("ForeColor", ForeColorDefault)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
  Cursor.MousePointer = PropBag.ReadProperty("MousePointer", 0)
  Cursor.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  'draw control
  ControlOK = True
  If UserControl.ScaleWidth < 20 Then ControlOK = False
  If UserControl.ScaleHeight < 20 Then ControlOK = False
  Call DrawControl
End Sub

Private Sub UserControl_Resize()
  If mValue <> 0 Then ControlOK = True
  If UserControl.ScaleWidth < 20 Then ControlOK = False
  If UserControl.ScaleHeight < 20 Then ControlOK = False
  Call DrawControl
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

  Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
  Call PropBag.WriteProperty("Value", mValue, ValueDefault)
  Call PropBag.WriteProperty("Positions", mPositions, PositionsDefault)
  Call PropBag.WriteProperty("ForeColor", mForeColor, ForeColorDefault)
  Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
  Call PropBag.WriteProperty("MousePointer", Cursor.MousePointer, 0)
  Call PropBag.WriteProperty("ToolTipText", Cursor.ToolTipText, "")
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=1,0,0,10
Public Property Get Positions() As Byte
  Positions = mPositions
End Property

Public Property Let Positions(ByVal New_Positions As Byte)
  If New_Positions > 0 Then mPositions = New_Positions
  PropertyChanged "Positions"
  Call DrawControl
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    mForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Cursor,Cursor,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = Cursor.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set Cursor.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Cursor,Cursor,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Cursor.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    Cursor.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Cursor,Cursor,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = Cursor.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    Cursor.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

