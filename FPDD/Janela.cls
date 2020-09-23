VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Janela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Parede"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarPosition As Long
Private mvarTamanho As Long
Private mvarParedeNum As Long
Private mvarSuporte As Parede

'----------------------------------------------- PUBLIC
Public Function DistanciaPonto(ByVal X As Long, ByVal Y As Long) As Long
  Dim Jx As Single
  Dim Jy As Single
  
  Call GetPos(Jx, Jy)
  DistanciaPonto = Hipotenusa(Jx - X, Jy - Y)
End Function

Public Function GotGrip(ByVal X As Long, ByVal Y As Long) As Boolean
  Dim hipo As Single
  Dim Res As Boolean
  Dim pX As Single
  Dim pY As Single
  
  Res = False
  Call GetPos(pX, pY)
  hipo = Hipotenusa(pX - X, pY - Y)
  If hipo <= mvarSuporte.Largura Then Res = True
  GotGrip = Res
End Function

Public Sub DrawJanela(ByRef Desenho As FloorPlan, ByVal Selected As Boolean)
  Dim MeioX As Integer, MeioY As Integer
  Dim Big As Integer
  Dim pX As Single, pY As Single
  Dim Pts(1 To 4) As Point
  Dim Tam2L As Single
  Dim Tam2J As Single
  Dim X1 As Long, Y1 As Long
  Dim X2 As Long, Y2 As Long
  Dim CentroX As Single, CentroY As Single
  
  If mvarSuporte.Visible = False Then Exit Sub
  Call GetPos(CentroX, CentroY)
  X1 = mvarSuporte.X1:  X2 = mvarSuporte.X2
  Y1 = mvarSuporte.Y1:  Y2 = mvarSuporte.Y2
  'Pontos - X
  Tam2J = Abs(Regra3(mvarSuporte.Tamanho, X2 - X1, mvarTamanho / 2, 0, 4)) / Desenho.PixCTM
  Tam2L = Abs(Regra3(mvarSuporte.Tamanho, Y2 - Y1, mvarSuporte.Largura / 2, 0, 4)) / Desenho.PixCTM 'rotação 90º
  pX = Desenho.MapToScreen(CentroX, 1)
  If Y2 > Y1 Then
    If X2 > X1 Then
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
    If X2 > X1 Then
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
  'Pontos - Y
  Tam2J = Abs(Regra3(mvarSuporte.Tamanho, Y2 - Y1, mvarTamanho / 2, 0, 4)) / Desenho.PixCTM
  Tam2L = Abs(Regra3(mvarSuporte.Tamanho, X2 - X1, mvarSuporte.Largura / 2, 0, 4)) / Desenho.PixCTM 'rotação 90º
  pY = Desenho.MapToScreen(CentroY, 2)
  If Y2 - Y1 > 0 Then
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
  Desenho.PicBox.FillStyle = 0
  Desenho.PicBox.FillColor = RGB(200, 220, 255)
  Desenho.PicBox.ForeColor = RGB(155, 155, 255)
  Call Polygon(Desenho.PicBox.hDC, Pts(1), 4)
  'desenhar indicação de janela seleccionada
  MeioX = Desenho.MapToScreen(CentroX, 1)
  MeioY = Desenho.MapToScreen(CentroY, 2)
  If Selected = True Then
    Big = mvarSuporte.Largura * Desenho.MinPTS \ Desenho.MinCTM
    If Big = 0 Then Big = 1
    Desenho.PicBox.DrawWidth = Big
    Desenho.PicBox.ForeColor = RGB(255, 0, 0)
    Desenho.PicBox.Line (MeioX, MeioY)-Step(0, 0)
    Desenho.PicBox.DrawWidth = 1
  End If
End Sub

Public Sub GetPos(ByRef X As Single, ByRef Y As Single)
  Dim X1 As Long
  Dim Y1 As Long
  Dim X2 As Long
  Dim Y2 As Long
  Dim MeioX As Single
  Dim MeioY As Single
  Dim TamP As Single
  
  'obter e tratar posições
  X1 = mvarSuporte.X1:  Y1 = mvarSuporte.Y1
  X2 = mvarSuporte.X2:  Y2 = mvarSuporte.Y2
  'obter centro e tamanhos
  TamP = mvarSuporte.Tamanho
  MeioX = Regra3(TamP, X2 - X1, mvarPosition + mvarTamanho \ 2, 0, 4)
  MeioY = Regra3(TamP, Y2 - Y1, mvarPosition + mvarTamanho \ 2, 0, 4)
  'devolver o valor correcto
  X = X1 + MeioX
  Y = Y1 + MeioY
End Sub

Public Sub ChangePos(ByRef Desenho As FloorPlan, ByVal X As Long, ByVal Y As Long)
  Dim Tmp As Long
  Dim RotX As Long, RotY As Long
  Dim Grau As Single
  Dim NewP As Long
  Dim TamP As Long
  Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
  
  'vector parede
  X1 = mvarSuporte.X1:  Y1 = mvarSuporte.Y1
  X2 = mvarSuporte.X2:  Y2 = mvarSuporte.Y2
  'passar coordenadas para centímetros
  X = ScreenToPlan(X, Desenho, 1)
  Y = ScreenToPlan(Y, Desenho, 2)
  'Calcula nova posição
  TamP = mvarSuporte.Tamanho
  NewP = Hipotenusa(X - X1, Y - Y1)
  NewP = NewP - mvarTamanho / 2
  'não sai dos limites da parede
  If NewP < 0 Then NewP = 0
  TamP = mvarSuporte.Tamanho
  If (NewP + mvarTamanho) > TamP Then
    NewP = TamP - mvarTamanho
  End If
  'set nova posição
  mvarPosition = NewP
End Sub

'----------------------------------------------- PROPERTIES
Public Property Set Suporte(ByRef wall As Parede)
  Set mvarSuporte = wall
End Property

Public Property Get Suporte() As Parede
  Set Suporte = mvarSuporte
End Property

Public Property Let ParedeNum(ByVal vData As Long)
    mvarParedeNum = vData
End Property

Public Property Get ParedeNum() As Long
    ParedeNum = mvarParedeNum
End Property

Public Property Let Tamanho(ByVal vData As Long)
    mvarTamanho = vData
End Property

Public Property Get Tamanho() As Long
    Tamanho = mvarTamanho
End Property

Public Property Let Position(ByVal vData As Long)
    mvarPosition = vData
End Property

Public Property Get Position() As Long
    Position = mvarPosition
End Property

'----------------------------------------------- PRIVATE
Private Function Hipotenusa(ByVal Cat1 As Long, ByVal Cat2 As Long) As Single
  Hipotenusa = Sqr(Cat1 * Cat1 + Cat2 * Cat2)
End Function

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

Private Function ScreenToPlan(ByVal Pos As Long, ByRef Desenho As FloorPlan, ByVal modo As Byte) As Long
  Dim Valor As Single
  
  If modo = 1 Then
    Valor = Pos - Desenho.TamanhoX / 2
    Valor = Valor * Desenho.PixCTM
    Valor = Valor + Desenho.MeioX - Desenho.CentroX
  Else
    Valor = Pos - Desenho.TamanhoY / 2
    Valor = Valor * Desenho.PixCTM
    Valor = Valor + Desenho.MeioY - Desenho.CentroY
  End If
  ScreenToPlan = Valor
End Function

