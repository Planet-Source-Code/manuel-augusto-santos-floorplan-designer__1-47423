VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Floor Plan Digital Designer - Paredes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'----------------------------------------------- PAREDE
Private mvarX1 As Long
Private mvarY1 As Long
Private mvarX2 As Long
Private mvarY2 As Long
Private oldX1 As Long
Private oldY1 As Long
Private oldX2 As Long
Private oldY2 As Long
Private oldPos As Boolean
Private mvarLargura As Integer
Private mvarVisible As Boolean
Private Pts(1 To 4) As Point

Private Type Point
  X As Long
  Y As Long
End Type

'----------------------------------------------- PUBLIC
Public Sub GetScreenPontos(ByRef ptsX() As Long, ByRef ptsY() As Long)
  ptsX(1) = Pts(1).X
  ptsY(1) = Pts(1).Y
  ptsX(2) = Pts(2).X
  ptsY(2) = Pts(2).Y
  ptsX(3) = Pts(3).X
  ptsY(3) = Pts(3).Y
  ptsX(4) = Pts(4).X
  ptsY(4) = Pts(4).Y
End Sub

Public Sub TrocaCoord()
  Dim Tmp As Long
  
  Tmp = mvarX1: mvarX1 = mvarX2: mvarX2 = Tmp
  Tmp = mvarY1: mvarY1 = mvarY2: mvarY2 = Tmp
End Sub

Public Function Tamanho() As Single
  Tamanho = Hipotenusa(mvarX2 - mvarX1, mvarY2 - mvarY1)
End Function

Public Sub GetMeio(ByRef X As Long, ByRef Y As Long)
  X = mvarX1 + (mvarX2 - mvarX1) / 2
  Y = mvarY1 + (mvarY2 - mvarY1) / 2
End Sub

Public Function GetTextX() As String
  Dim Res As String
  
  If (mvarX2 - mvarX1 >= 0) Then
    Res = "X:+" & (mvarX2 - mvarX1)
   Else
    Res = "X:-" & (mvarX1 - mvarX2)
  End If
  GetTextX = Res
End Function

Public Function GetTextY() As String
  Dim Res As String
  
  If (mvarY2 - mvarY1 >= 0) Then
    Res = "Y:+" & (mvarY2 - mvarY1)
   Else
    Res = "Y:-" & (mvarY1 - mvarY2)
  End If
  GetTextY = Res
End Function

Public Function GotGrip(ByVal X As Long, ByVal Y As Long) As Byte
  Dim hipo As Single
  Dim Res As Byte
  
  Res = 0
  hipo = Hipotenusa(mvarX1 - X, mvarY1 - Y)
  If hipo <= mvarLargura Then
    Res = 1
  Else
    hipo = Hipotenusa(mvarX2 - X, mvarY2 - Y)
    If hipo <= mvarLargura Then Res = 2
  End If
  GotGrip = Res
End Function

Public Function DistanciaPonto(ByVal X As Long, ByVal Y As Long) As Long
  Dim A As Long, B As Long, C As Long, D As Long
  Dim pX1 As Long, pY1 As Long
  Dim pX2 As Long, pY2 As Long
  Dim Res As Single
  Dim RotX As Long, RotY As Long
  Dim Grau As Single
  
  pX1 = mvarX1: pY1 = mvarY1
  pX2 = mvarX2: pY2 = mvarY2
  Call NormalizeCoord(pX1, pY1, pX2, pY2)
  A = X - pX1
  B = Y - pY1
  C = pX2 - pX1
  D = pY2 - pY1
  If (C <> 0) Or (D <> 0) Then
    If Abs(C) >= Abs(D) Then
      Grau = -Atn(D / C)
      Call Rotate2D(0, 0, A, B, Grau, RotX, RotY)
      RotY = -RotY
    ElseIf Abs(D) >= Abs(C) Then
      If (pY1 > pY2) Then
        Grau = 1.570796 - Abs(Atn(C / D))
        Call Rotate2D(0, 0, A, B, Grau, RotX, RotY)
        RotY = -RotY
      Else
        Grau = 1.570796 + Abs(Atn(C / D))
        Call Rotate2D(0, 0, A, B, Grau, RotX, RotY)
        RotX = -RotX
      End If
    End If
    If RotX < 0 Then
      Res = Hipotenusa(X - pX1, Y - pY1)
    ElseIf RotX > Tamanho Then
      Res = Hipotenusa(X - pX2, Y - pY2)
    Else
      Res = Abs(A * D - B * C) / Hipotenusa(C, D)
    End If
  Else
    Res = Hipotenusa(A, B)
  End If
  DistanciaPonto = Res
End Function

Public Sub DrawParede(ByRef Plan As FloorPlan, ByVal Selected As Boolean, ByVal ForeColor As Long, ByVal DrawOldPos As Boolean)
Attribute DrawParede.VB_Description = "Desenhar Parede"
Attribute DrawParede.VB_UserMemId = 0
  Dim pX1 As Long, pY1 As Long
  Dim pX2 As Long, pY2 As Long
  Dim pX As Single, pY As Single
  Dim Tam2L As Single
  Dim Tam2J As Single
  Dim Big As Integer
  Dim DrawWall As Boolean
  
  If mvarVisible = False Then Exit Sub
  Big = (mvarLargura * Plan.MinPTS \ Plan.MinCTM) / 2
  If Big = 0 Then Big = 1
  Plan.PicBox.DrawWidth = 1
  Plan.PicBox.ForeColor = ForeColor
  DrawWall = False
  If DrawOldPos = False Then
    pX1 = MapToScreen(1, mvarX1, Plan)
    pX2 = MapToScreen(1, mvarX2, Plan)
    pY1 = MapToScreen(2, mvarY1, Plan)
    pY2 = MapToScreen(2, mvarY2, Plan)
    oldX1 = mvarX1: oldX2 = mvarX2
    oldY1 = mvarY1: oldY2 = mvarY2
    oldPos = True
    DrawWall = True
  Else
    If oldPos Then
      pX1 = MapToScreen(1, oldX1, Plan)
      pX2 = MapToScreen(1, oldX2, Plan)
      pY1 = MapToScreen(2, oldY1, Plan)
      pY2 = MapToScreen(2, oldY2, Plan)
      oldPos = False
      DrawWall = True
    End If
  End If
  If (DrawWall = True) And (Tamanho > 0) Then
    'determinar ponto medio
    pX = pX1 + (pX2 - pX1) / 2
    pY = pY1 + (pY2 - pY1) / 2
    'calcular dimensões (com rotação) do polígono - parte X
    Tam2J = Abs((mvarX2 - mvarX1) / 2) / Plan.PixCTM
    Tam2L = Abs(Regra3(Tamanho, mvarY2 - mvarY1, mvarLargura / 2, 0, 4)) / Plan.PixCTM 'rotação 90º
    'determinar pontos - X
    If mvarY2 > mvarY1 Then
      If mvarX2 > mvarX1 Then
        Pts(1).X = pX - Tam2J + Tam2L
        Pts(2).X = pX + Tam2J + Tam2L
        Pts(3).X = pX + Tam2J - Tam2L
        Pts(4).X = pX - Tam2J - Tam2L
      Else
        Pts(1).X = pX + Tam2J - Tam2L
        Pts(2).X = pX - Tam2J - Tam2L
        Pts(3).X = pX - Tam2J + Tam2L
        Pts(4).X = pX + Tam2J + Tam2L
      End If
    Else
      If mvarX2 > mvarX1 Then
        Pts(1).X = pX - Tam2J - Tam2L
        Pts(2).X = pX + Tam2J - Tam2L
        Pts(3).X = pX + Tam2J + Tam2L
        Pts(4).X = pX - Tam2J + Tam2L
      Else
        Pts(1).X = pX + Tam2J + Tam2L
        Pts(2).X = pX - Tam2J + Tam2L
        Pts(3).X = pX - Tam2J - Tam2L
        Pts(4).X = pX + Tam2J - Tam2L
      End If
    End If
    'calcular dimensões (com rotação) do polígono - parte X
    Tam2J = Abs((mvarY2 - mvarY1) / 2) / Plan.PixCTM
    Tam2L = Abs(Regra3(Tamanho, mvarX2 - mvarX1, mvarLargura / 2, 0, 4)) / Plan.PixCTM 'rotação 90º
    'determinar pontos - Y
    If mvarY2 - mvarY1 > 0 Then
      Pts(1).Y = pY - Tam2J - Tam2L
      Pts(2).Y = pY + Tam2J - Tam2L
      Pts(3).Y = pY + Tam2J + Tam2L
      Pts(4).Y = pY - Tam2J + Tam2L
    Else
      Pts(1).Y = pY + Tam2J - Tam2L
      Pts(2).Y = pY - Tam2J - Tam2L
      Pts(3).Y = pY - Tam2J + Tam2L
      Pts(4).Y = pY + Tam2J + Tam2L
    End If
    'desenhar Polígono
    Plan.PicBox.FillStyle = 0
    If Selected = True Then
      Plan.PicBox.FillColor = RGB(255, 0, 0)
      Plan.PicBox.ForeColor = RGB(255, 0, 0)
      Plan.PicBox.DrawWidth = Big
      Plan.PicBox.Line (pX1, pY1)-Step(0, 0)
      Plan.PicBox.Line (pX2, pY2)-Step(0, 0)
      Plan.PicBox.DrawWidth = 1
      Plan.PicBox.ForeColor = ForeColor
    Else
      Plan.PicBox.FillColor = ForeColor
      Plan.PicBox.ForeColor = ForeColor
    End If
    Call Polygon(Plan.PicBox.hDC, Pts(1), 4)
  End If
  Plan.PicBox.DrawWidth = 1
  Plan.PicBox.ForeColor = 0
End Sub

Public Property Let Largura(ByVal vData As Integer)
Attribute Largura.VB_Description = "Largura da Parede em Cm"
    mvarLargura = vData
End Property

Public Property Get Largura() As Integer
    Largura = mvarLargura
End Property

Public Property Get Y2() As Long
Attribute Y2.VB_Description = "Posição da Parede - Read Only"
    Y2 = mvarY2
End Property

Public Property Get X2() As Long
Attribute X2.VB_Description = "Posição da Parede - Read Only"
    X2 = mvarX2
End Property

Public Property Get Y1() As Long
Attribute Y1.VB_Description = "Posição da Parede - Read Only"
    Y1 = mvarY1
End Property

Public Property Get X1() As Long
Attribute X1.VB_Description = "Posição da parede - Read Only"
    X1 = mvarX1
End Property

Public Property Let Y2(ByVal vData As Long)
    mvarY2 = vData
End Property

Public Property Let X2(ByVal vData As Long)
    mvarX2 = vData
End Property

Public Property Let Y1(ByVal vData As Long)
    mvarY1 = vData
End Property

Public Property Let X1(ByVal vData As Long)
    mvarX1 = vData
End Property

Public Property Let Visible(ByVal vData As Boolean)
    mvarVisible = vData
    If mvarVisible = False Then oldPos = False
End Property

Public Property Get Visible() As Boolean
    Visible = mvarVisible
End Property

'----------------------------------------------- PRIVATE
Private Function MapToScreen(ByVal modo As Byte, ByVal Pos As Single, ByRef Plan As FloorPlan) As Long
  Dim Valor As Single
  Dim Aux As Single
  
  If modo = 1 Then 'modo X
    Valor = Pos + Plan.CentroX - Plan.MeioX
    Valor = Valor * Plan.MinPTS
    Aux = Plan.TamanhoX / 2
  Else             'modo Y
    Valor = Pos + Plan.CentroY - Plan.MeioY
    Valor = Valor * Plan.MinPTS
    Aux = Plan.TamanhoY / 2
  End If
  Valor = Valor / Plan.MinCTM + Aux
  MapToScreen = Round(Valor, 0)
End Function

Private Sub Class_Initialize()
  mvarX1 = 0: mvarY1 = 0
  mvarX2 = 0: mvarY2 = 0
  oldX1 = 0:  oldY1 = 0
  oldX2 = 0:  oldY2 = 0
  oldPos = False
  mvarVisible = True
  mvarLargura = 1
End Sub

Private Function Hipotenusa(ByVal Cat1 As Long, ByVal Cat2 As Long) As Single
  Hipotenusa = Sqr(Cat1 * Cat1 + Cat2 * Cat2)
End Function

Private Sub NormalizeCoord(ByRef X1 As Long, ByRef Y1 As Long, ByRef X2 As Long, ByRef Y2 As Long)
  Dim Tmp As Long
  
  If X1 > X2 Then
    Tmp = X2: X2 = X1: X1 = Tmp
    Tmp = Y2: Y2 = Y1: Y1 = Tmp
  ElseIf X1 = X2 Then
    If Y1 > Y2 Then
      Tmp = X2: X2 = X1: X1 = Tmp
      Tmp = Y2: Y2 = Y1: Y1 = Tmp
    End If
  End If
End Sub

Private Sub Rotate2D(ByVal PosX As Long, ByVal PosY As Long, _
                    ByVal DirX As Long, ByVal DirY As Long, _
                    ByVal Grau As Single, _
                    ByRef RotX As Long, ByRef RotY As Long)
  Dim RX As Single
  Dim RY As Single
  
  RX = PosX + (DirX - PosX) * Cos(Grau) - (DirY - PosY) * Sin(Grau)
  RY = PosY + (DirX - PosX) * Sin(Grau) + (DirY - PosY) * Cos(Grau)
  RotX = Round(RX, 0)
  RotY = Round(RY, 0)
End Sub

Private Function Regra3(ByVal N11 As Single, ByVal N12 As Single, ByVal N21 As Single, ByVal N22 As Single, ByVal modo As Byte) As Single
  Dim Res As Single
  
  Select Case modo
    Case 1: Res = (N12 * N21) / N22
    Case 2: Res = (N11 * N22) / N21
    Case 3: Res = (N11 * N22) / N12
    Case 4: Res = (N12 * N21) / N11
  End Select
  Regra3 = Res
End Function


