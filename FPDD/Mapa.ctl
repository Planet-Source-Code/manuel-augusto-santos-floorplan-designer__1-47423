VERSION 5.00
Begin VB.UserControl Mapa 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   ScaleHeight     =   100
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   100
   ToolboxBitmap   =   "Mapa.ctx":0000
   Begin VB.Timer DragTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   90
   End
   Begin VB.PictureBox ECR 
      Height          =   420
      Left            =   360
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   0
      Top             =   630
      Width           =   735
      Begin VB.Image ECRImage 
         Height          =   360
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   675
      End
   End
End
Attribute VB_Name = "Mapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Windows API Functions
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
'Private properties
Private mMaxX As Long        'dimensão X do mapa (scalewidth)
Private mMaxY As Long        'dimensão Y do mapa (scaleheight)
Private DragMapa As Boolean  'o utilizador está a puxar o mapa
Private PosDragX As Long     'posição inicial X do puxar o mapa
Private PosDragY As Long     'posição inicial Y do puxar o mapa
Private oldPosX As Long      'posição antiga X para não andar a actualizar
Private oldPosY As Long      'posição antiga Y para não andar a actualizar
'Default Property Values:
Const EscalaDefault = 1
Const MaxXdefault = 100
Const MaxYdefault = 100
'Property Variables:
Dim mEscala As Long          'Escala para o mapa (1 pixel = x units reais)
'Event Declarations:
Event ChangePos(ByVal X As Long, ByVal Y As Long)

'----------------------------------------------- PUBLIC
Public Sub DrawMapa(ByVal cmX As Long, ByVal cmY As Long, ByVal cmTamX As Long, ByVal cmTamY As Long)
  Dim Mapa1X As Long
  Dim Mapa1Y As Long
  Dim Mapa2X As Long
  Dim Mapa2Y As Long
  
  'calcular posição da janelinha
  Mapa1X = (cmX - cmTamX / 2) / mEscala
  Mapa1Y = (cmY - cmTamY / 2) / mEscala
  Mapa2X = (cmX + cmTamX / 2) / mEscala
  Mapa2Y = (cmY + cmTamY / 2) / mEscala
  UserControl.ECR.top = Mapa1Y - 2
  UserControl.ECR.Left = Mapa1X - 2
  UserControl.ECR.Width = Mapa2X - Mapa1X + 1
  UserControl.ECR.Height = Mapa2Y - Mapa1Y + 1
  DoEvents
End Sub

'----------------------------------------------- PRIVATE
Private Sub DragTimer_Timer()
  Dim Mpos As Point
  Dim Wpos As Point
  Dim PosX As Long
  Dim PosY As Long
  Dim MeioX As Long
  Dim MeioY As Long
  
  If DragMapa = False Then DragTimer.Enabled = False
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Call ClientToScreen(UserControl.hWnd, Wpos)
  PosX = Mpos.X - Wpos.X - PosDragX
  PosY = Mpos.Y - Wpos.Y - PosDragY
  If (PosX + UserControl.ECR.Width / 2 < 0) Then PosX = -UserControl.ECR.Width / 2
  If (PosY + UserControl.ECR.Height / 2 < 0) Then PosY = -UserControl.ECR.Height / 2
  If (PosX + UserControl.ECR.Width / 2 > mMaxX) Then PosX = mMaxX - UserControl.ECR.Width / 2
  If (PosY + UserControl.ECR.Height / 2 > mMaxY) Then PosY = mMaxY - UserControl.ECR.Height / 2
  If UserControl.ECR.BorderStyle = 1 Then
    UserControl.ECR.top = PosY - 2
    UserControl.ECR.Left = PosX - 2
    PosX = UserControl.ECR.Left + UserControl.ECR.Width / 2 + 2
    PosY = UserControl.ECR.top + UserControl.ECR.Height / 2 + 2
  Else
    UserControl.ECR.top = PosY
    UserControl.ECR.Left = PosX
    PosX = 2 + UserControl.ECR.Left + UserControl.ECR.Width / 2
    PosY = 2 + UserControl.ECR.top + UserControl.ECR.Height / 2
  End If
  DoEvents
  PosX = PosX * mEscala
  PosY = PosY * mEscala
  If (oldPosX <> PosX) Or (oldPosY <> PosY) Then RaiseEvent ChangePos(PosX, PosY)
  oldPosX = PosX
  oldPosY = PosY
End Sub

Private Sub ECR_Resize()
  ECRImage.top = 0
  ECRImage.Left = 0
  ECRImage.Width = ECR.Width
  ECRImage.Height = ECR.Height
End Sub

Private Sub ECRImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Mpos As Point
  Dim Wpos As Point
  
  DragMapa = True
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Call ClientToScreen(UserControl.ECR.hWnd, Wpos)
  PosDragX = Mpos.X - Wpos.X
  PosDragY = Mpos.Y - Wpos.Y
  DragTimer.Enabled = True
End Sub

Private Sub ECRImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Mpos As Point
  Dim Wpos As Point
  
  DragMapa = False
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Call ClientToScreen(UserControl.ECR.hWnd, Wpos)
  PosDragX = Mpos.X - Wpos.X
  PosDragY = Mpos.Y - Wpos.Y
  DragTimer.Enabled = False
  Call DragTimer_Timer
End Sub

'----------------------------------------------- Properties
'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    mEscala = EscalaDefault
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set UserControl.Picture = PropBag.ReadProperty("BackPicture", Nothing)
    UserControl.BorderStyle = PropBag.ReadProperty("BackBorder", 1)
    mEscala = PropBag.ReadProperty("Escala", EscalaDefault)
    Set ECRImage.Picture = PropBag.ReadProperty("ScrPicture", Nothing)
    ECR.BorderStyle = PropBag.ReadProperty("ScrBorder", 1)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    ECRImage.MousePointer = PropBag.ReadProperty("MousePointer", 0)
End Sub

Private Sub UserControl_Resize()
    mMaxX = UserControl.ScaleWidth
    mMaxY = UserControl.ScaleHeight
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackPicture", UserControl.Picture, Nothing)
    Call PropBag.WriteProperty("BackBorder", UserControl.BorderStyle, 1)
    Call PropBag.WriteProperty("Escala", mEscala, EscalaDefault)
    Call PropBag.WriteProperty("ScrPicture", ECRImage.Picture, Nothing)
    Call PropBag.WriteProperty("ScrBorder", ECR.BorderStyle, 1)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", ECRImage.MousePointer, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get BackPicture() As Picture
Attribute BackPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set BackPicture = UserControl.Picture
End Property

Public Property Set BackPicture(ByVal New_BackPicture As Picture)
    Set UserControl.Picture = New_BackPicture
    PropertyChanged "BackPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BackBorder() As MSComctlLib.BorderStyleConstants
Attribute BackBorder.VB_Description = "Returns/sets the border style for an object."
    BackBorder = UserControl.BorderStyle
End Property

Public Property Let BackBorder(ByVal New_BackBorder As MSComctlLib.BorderStyleConstants)
    UserControl.BorderStyle() = New_BackBorder
    PropertyChanged "BackBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,1
Public Property Get Escala() As Long
    Escala = mEscala
End Property

Public Property Let Escala(ByVal New_Escala As Long)
    mEscala = New_Escala
    PropertyChanged "Escala"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ECR,ECR,-1,BorderStyle
Public Property Get ScrBorder() As MSComctlLib.BorderStyleConstants
Attribute ScrBorder.VB_Description = "Returns/sets the border style for an object."
    ScrBorder = ECR.BorderStyle
End Property

Public Property Let ScrBorder(ByVal New_ScrBorder As MSComctlLib.BorderStyleConstants)
    ECR.BorderStyle() = New_ScrBorder
    PropertyChanged "ScrBorder"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ECRImage,ECRImage,-1,Picture
Public Property Get ScrPicture() As Picture
Attribute ScrPicture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set ScrPicture = ECRImage.Picture
End Property

Public Property Set ScrPicture(ByVal New_ScrPicture As Picture)
    Set ECRImage.Picture = New_ScrPicture
    PropertyChanged "ScrPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ECRImage,ECRImage,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = ECRImage.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set ECRImage.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ECRImage,ECRImage,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = ECRImage.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ECRImage.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

