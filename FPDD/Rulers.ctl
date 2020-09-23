VERSION 5.00
Begin VB.UserControl Rulers 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00ACFBFF&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   ForeColor       =   &H00000000&
   MouseIcon       =   "Rulers.ctx":0000
   MousePointer    =   99  'Custom
   PropertyPages   =   "Rulers.ctx":0152
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   294
   ToolboxBitmap   =   "Rulers.ctx":0168
   Begin VB.Timer DragTimer 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   90
      Top             =   0
   End
End
Attribute VB_Name = "Rulers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'Windows API Functions
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
'Private
Private mMinCTM(11) As Integer 'medida mínima em centímetros
Private mMinPTS(11) As Integer 'número mínimo de pontos por unidade
Private mActive As Boolean     'Ruler pronto (com escala)
Private CursorPos As Long      'Pixel da linha indicadora de posição
Private OldCursor As Long      'Pixel da posição antiga de CursorPos
Private CursorON As Boolean    'Linha visível ou invisível
Private DragRuler As Boolean   'o utilizador está a puxar o ruler
Private PosStartDrag As Long   'posição inicial do «puxar o ruler»
Private mMeio As Long          'posição (cm) do meio do ruler
Private mZoom As Integer       'número da escala de valores
'Property Variables:
Dim m_Orientation As MSComctlLib.OrientationConstants
Private mMinValue As Long      'Mínimo valor possível
Private mMaxValue As Long      'Máximo valor possível
Private mNums As Picture       'bitmap com o desenho dos números
Private mOrientation As MSComctlLib.OrientationConstants '0=Horizontal, 1=Vertical
'Default Property Values:
Const MinValueDefault = 0
Const MaxValueDefault = 0
Const ActiveDefault = False
Const OrientationDefault = 0
'Event Declarations:
Event ChangePos(NovoMeio As Long)
'Types
Private Type Point
  X As Long
  Y As Long
End Type


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub HideCursor()
   UserControl.DrawMode = 6
   If (mOrientation = ccOrientationHorizontal) And (CursorON = True) Then
     UserControl.Line (OldCursor, 0)-Step(0, UserControl.ScaleHeight)
     CursorON = False
   End If
   If (mOrientation = ccOrientationVertical) And (CursorON = True) Then
     UserControl.Line (0, OldCursor)-Step(UserControl.ScaleWidth, 0)
     CursorON = False
   End If
   UserControl.DrawMode = 13
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub SetCursor(ByVal Pos As Long)
   UserControl.DrawMode = 6
   CursorPos = Pos
   If mOrientation = ccOrientationHorizontal Then
     If CursorON Then
       UserControl.Line (OldCursor, 0)-Step(0, UserControl.ScaleHeight)
     End If
     UserControl.Line (CursorPos, 0)-Step(0, UserControl.ScaleHeight)
     OldCursor = CursorPos
     CursorON = True
   Else
     If CursorON Then
       UserControl.Line (0, OldCursor)-Step(UserControl.ScaleWidth, 0)
     End If
     UserControl.Line (0, Pos)-Step(UserControl.ScaleWidth, 0)
     OldCursor = CursorPos
     CursorON = True
   End If
   UserControl.DrawMode = 13
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=5
Public Sub Draw(ByVal Escala As Integer, ByVal MeioCM As Long, ByVal CentroCM As Long, ByVal Grid As Boolean, Optional ByVal Tamanho As Long, Optional ByRef Plan As PictureBox)
  Dim PixelP As Integer
  Dim Tam As Byte
  Dim Pos As Integer
  Dim CorAux As Long
    
  If Active = False Then Exit Sub
  mMeio = MeioCM
  mZoom = Escala
  If Grid Then CorAux = Plan.ForeColor
  UserControl.Cls
  CursorON = False
  'desenhar Vertical
  If mOrientation = ccOrientationVertical Then
    PixelP = UserControl.ScaleHeight
    If (PixelP And 1) = 1 Then PixelP = PixelP - 1
    PixelP = PixelP \ 2
    Pos = MeioCM
    Do
      Tam = GetTamRuler(Escala, Pos - CentroCM)
      UserControl.ForeColor = 0
      UserControl.Line (UserControl.ScaleWidth - Tam, PixelP)-Step(Tam, 0)
      If (Tam > UserControl.ScaleWidth \ 2) Then Call DrawNum((Pos - CentroCM) \ 10, 1, PixelP - 5)
      If Grid Then
        Plan.ForeColor = RGB(200, 200, 200)
        If (Tam > UserControl.ScaleWidth \ 2) Then Plan.Line (0, PixelP)-Step(Tamanho, 0)
      End If
      PixelP = PixelP + mMinPTS(Escala)
      Pos = Pos + mMinCTM(Escala)
    Loop Until PixelP > UserControl.ScaleHeight
    PixelP = UserControl.ScaleHeight
    If (PixelP And 1) = 1 Then PixelP = PixelP - 1
    PixelP = PixelP \ 2 - mMinPTS(Escala)
    Pos = MeioCM - mMinCTM(Escala)
    Do
      Tam = GetTamRuler(Escala, Pos - CentroCM)
      UserControl.ForeColor = 0
      UserControl.Line (UserControl.ScaleWidth - Tam, PixelP)-Step(Tam, 0)
      If (Tam > UserControl.ScaleWidth \ 2) Then Call DrawNum((Pos - CentroCM) \ 10, 1, PixelP - 5)
      If Grid Then
        Plan.ForeColor = RGB(200, 200, 200)
        If (Tam > UserControl.ScaleWidth \ 2) Then Plan.Line (0, PixelP)-Step(Tamanho, 0)
      End If
      PixelP = PixelP - mMinPTS(Escala)
      Pos = Pos - mMinCTM(Escala)
    Loop Until PixelP < 0
  End If
  'desenha o de cima
  If mOrientation = ccOrientationHorizontal Then
    PixelP = UserControl.ScaleWidth
    If (PixelP And 1) = 1 Then PixelP = PixelP - 1
    PixelP = PixelP \ 2
    Pos = MeioCM
    Do
      Tam = GetTamRuler(Escala, Pos - CentroCM)
      UserControl.ForeColor = 0
      UserControl.Line (PixelP, UserControl.ScaleHeight - Tam)-Step(0, Tam)
      If (Tam > UserControl.ScaleHeight \ 2) Then Call DrawNum((Pos - CentroCM) \ 10, PixelP + 1, 1)
      If Grid Then
        Plan.ForeColor = RGB(200, 200, 200)
        If (Tam > UserControl.ScaleHeight \ 2) Then Plan.Line (PixelP, 0)-Step(0, Tamanho)
      End If
      PixelP = PixelP + mMinPTS(Escala)
      Pos = Pos + mMinCTM(Escala)
    Loop Until PixelP > UserControl.ScaleWidth
    PixelP = UserControl.ScaleWidth
    If (PixelP And 1) = 1 Then PixelP = PixelP - 1
    PixelP = PixelP \ 2 - mMinPTS(Escala)
    Pos = MeioCM - mMinCTM(Escala)
    Do
      Tam = GetTamRuler(Escala, Pos - CentroCM)
      UserControl.ForeColor = 0
      UserControl.Line (PixelP, UserControl.ScaleHeight - Tam)-Step(0, Tam)
      If (Tam > UserControl.ScaleHeight \ 2) Then Call DrawNum((Pos - CentroCM) \ 10, PixelP + 1, 1)
      If Grid = True Then
        Plan.ForeColor = RGB(200, 200, 200)
        If (Tam > UserControl.ScaleHeight \ 2) Then Plan.Line (PixelP, 0)-Step(0, Tamanho)
      End If
      PixelP = PixelP - mMinPTS(Escala)
      Pos = Pos - mMinCTM(Escala)
    Loop Until PixelP < 0
  End If
  If Grid Then Plan.ForeColor = CorAux
End Sub

Public Sub SetScale(ByRef MinPTS() As Integer, ByRef MinCTM() As Integer)
  Dim i As Byte
  
  For i = 0 To 10
     mMinPTS(i) = MinPTS(i)
     mMinCTM(i) = MinCTM(i)
  Next i
End Sub

Private Function GetTamRuler(ByVal Escala As Byte, ByVal Valor As Long) As Byte
  Dim Big As Byte
  
  Big = GetImportance(Escala, Valor)
  If mOrientation = ccOrientationHorizontal Then
    Select Case Big
      Case 1: GetTamRuler = 0.25 * UserControl.ScaleHeight
      Case 2: GetTamRuler = 0.4 * UserControl.ScaleHeight
      Case 3: GetTamRuler = 0.6 * UserControl.ScaleHeight
      Case 4: GetTamRuler = UserControl.ScaleHeight - 5
    End Select
  Else
    Select Case Big
      Case 1: GetTamRuler = 0.25 * UserControl.ScaleWidth
      Case 2: GetTamRuler = 0.4 * UserControl.ScaleWidth
      Case 3: GetTamRuler = 0.6 * UserControl.ScaleWidth
      Case 4: GetTamRuler = UserControl.ScaleWidth - 5
    End Select
  End If
End Function

Private Function GetImportance(ByVal Escala As Byte, ByVal Valor As Long) As Byte
  Dim val2 As Long
  Dim Val3 As Long
  Dim val4 As Long
  
  val2 = mMinCTM(Escala) * 5
  Val3 = val2 * 2
  val4 = Val3 * 10
  If (Valor Mod val4) = 0 Then GetImportance = 4: Exit Function
  If (Valor Mod Val3) = 0 Then GetImportance = 3: Exit Function
  If (Valor Mod val2) = 0 Then GetImportance = 2: Exit Function
  GetImportance = 1
End Function

Private Sub DrawNum(ByVal Num As Long, ByVal X As Long, ByVal Y As Long)
  Dim texto As String
  Dim i As Integer
  Dim j As Integer

  texto = Str(Num)
  j = 0
  For i = Len(texto) To 1 Step -1
    If mOrientation = ccOrientationHorizontal Then
      Select Case Mid(texto, i, 1)
        Case "0": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 0, 0, 5, 5
        Case "1": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 5, 0, 5, 5
        Case "2": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 10, 0, 5, 5
        Case "3": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 15, 0, 5, 5
        Case "4": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 20, 0, 5, 5
        Case "5": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 25, 0, 5, 5
        Case "6": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 30, 0, 5, 5
        Case "7": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 35, 0, 5, 5
        Case "8": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 40, 0, 5, 5
        Case "9": UserControl.PaintPicture mNums, X - j, Y, 5, 5, 45, 0, 5, 5
      End Select
    Else
      Select Case Mid(texto, i, 1)
        Case "0": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 45, 5, 5
        Case "1": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 40, 5, 5
        Case "2": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 35, 5, 5
        Case "3": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 30, 5, 5
        Case "4": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 25, 5, 5
        Case "5": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 20, 5, 5
        Case "6": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 15, 5, 5
        Case "7": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 10, 5, 5
        Case "8": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 5, 5, 5
        Case "9": UserControl.PaintPicture mNums, X, Y + j, 5, 5, 0, 0, 5, 5
      End Select
    End If
    j = j + 5
  Next i
End Sub

Private Sub SetRulerMeio(Dif As Long)
  Dim NPos As Long
  
  NPos = mMeio - (Dif \ mMinPTS(mZoom)) * mMinCTM(mZoom)
  If NPos > mMaxValue Then NPos = mMaxValue
  If NPos < mMinValue Then NPos = mMinValue
  If NPos <> mMeio Then mMeio = NPos
End Sub

'=======================================================================

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,3,0,0
Public Property Get Numeros() As Picture
    If Ambient.UserMode Then Err.Raise 393
    Set Numeros = mNums
End Property

Public Property Set Numeros(ByVal New_Numeros As Picture)
    If Ambient.UserMode Then Err.Raise 382
    Set mNums = New_Numeros
    PropertyChanged "Numeros"
End Property

Private Sub DragTimer_Timer()
  Dim Mpos As Point
  Dim Wpos As Point
  Dim Pos As Long
  
  If DragRuler = False Then DragTimer.Enabled = False
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Call ClientToScreen(UserControl.hWnd, Wpos)
  If mOrientation = ccOrientationHorizontal Then Pos = Mpos.X - Wpos.X Else Pos = Mpos.Y - Wpos.Y
  SetRulerMeio Pos - PosStartDrag
  PosStartDrag = Pos
  RaiseEvent ChangePos(mMeio)
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    Set mNums = LoadPicture("")
    mOrientation = OrientationDefault
    mActive = ActiveDefault
    CursorON = False
    DragRuler = False
    mMinValue = MinValueDefault
    mMaxValue = MaxValueDefault
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mNums = PropBag.ReadProperty("Numeros", Nothing)
    mOrientation = PropBag.ReadProperty("Orientation", OrientationDefault)
    mMinValue = PropBag.ReadProperty("MinValue", MinValueDefault)
    mMaxValue = PropBag.ReadProperty("MaxValue", MaxValueDefault)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Numeros", mNums, Nothing)
    Call PropBag.WriteProperty("Orientation", mOrientation, OrientationDefault)
    Call PropBag.WriteProperty("MinValue", mMinValue, MinValueDefault)
    Call PropBag.WriteProperty("MaxValue", mMaxValue, MaxValueDefault)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,2,False
Public Property Get Active() As Boolean
Attribute Active.VB_MemberFlags = "400"
    Active = mActive
End Property

Public Property Let Active(ByVal New_Active As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    mActive = New_Active
End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mActive = False Then Exit Sub
    DragRuler = True
    If mOrientation = ccOrientationHorizontal Then PosStartDrag = X Else PosStartDrag = Y
    DragTimer.Enabled = True
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragRuler = False
    DragTimer.Enabled = False
    If mActive = False Then Exit Sub
    Call DragTimer_Timer
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MinValue() As Long
Attribute MinValue.VB_ProcData.VB_Invoke_Property = "Ruler"
    MinValue = mMinValue
End Property

Public Property Let MinValue(ByVal New_MinValue As Long)
    mMinValue = New_MinValue
    PropertyChanged "MinValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get MaxValue() As Long
Attribute MaxValue.VB_ProcData.VB_Invoke_Property = "Ruler"
    MaxValue = mMaxValue
End Property

Public Property Let MaxValue(ByVal New_MaxValue As Long)
    mMaxValue = New_MaxValue
    PropertyChanged "MaxValue"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,3,0,0
Public Property Get Orientation() As MSComctlLib.OrientationConstants
    If Ambient.UserMode Then Err.Raise 393
    Orientation = mOrientation
End Property

Public Property Let Orientation(ByVal New_Orientation As MSComctlLib.OrientationConstants)
    If Ambient.UserMode Then Err.Raise 382
    mOrientation = New_Orientation
    PropertyChanged "Orientation"
End Property

