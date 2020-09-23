Attribute VB_Name = "MBTools"
Option Explicit
Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoints As Any, ByVal nCount As Long) As Long
'----------------------------------------------------------
Public PolyG(200, 50) As Point '\
Public PolyMax(200) As Integer ' | definição dos polígonos
Public Polygons As Integer     '/
Public PolyCF(200) As Long '\
Public PolyCD(200) As Long ' | definição do aspecto dos polígonos
Public PolyD(200) As Byte  '/
Public Dbox(12) As Integer ' patterns
'----------------------------------------------------------
Public Paredes As New Collection
Public Janelas As New Collection
Public Desenho As New FloorPlan
Public SelectP As Integer ' número da parede seleccionada
Public SelectJ As Integer ' número da janela seleccionada

'+--------------------------------------------------------+
'| Tratamento de Junções de paredes                       |
'+--------------------------------------------------------+
Public Function Intercept(L1X1 As Long, L1Y1 As Long, L1X2 As Long, L1Y2 As Long, _
                          L2X1 As Long, L2Y1 As Long, L2X2 As Long, L2Y2 As Long) As Point
  Dim A As Long, B As Long, C As Long, D As Long, E As Long, F As Long
  Dim L As Single
  
  Intercept.X = 0:  Intercept.Y = 0
  A = L1X2 - L1X1:  B = L1Y2 - L1Y1
  C = L2X2 - L2X1:  D = L2Y2 - L2Y1
  E = L2X1 - L1X1:  F = L2Y1 - L1Y1
  If (A * D = B * C) Then Exit Function
  L = (A * F - B * E) / (B * C - A * D)
  Intercept.X = L2X1 + L * C
  Intercept.Y = L2Y1 + L * D
End Function

Public Sub Junction(ByRef P1 As Parede, ByRef P2 As Parede)
  Dim A As Long, B As Long, C As Long, D As Long
  Dim pts1X(1 To 4) As Long, pts1Y(1 To 4) As Long
  Dim pts2X(1 To 4) As Long, pts2Y(1 To 4) As Long
  Dim pts(1 To 4) As Point
    
  'se forem paralelos não faz (angulo <10º)
  A = P1.X2 - P1.X1:  B = P1.Y2 - P1.Y1
  C = P2.X2 - P1.X1:  D = P2.Y2 - P2.Y1
  If Abs(A * D - B * C) < 1 Then Exit Sub
  'obtem pontos das paredes
  P1.GetScreenPontos pts1X, pts1Y
  P2.GetScreenPontos pts2X, pts2Y
  pts(1) = Intercept(pts1X(1), pts1Y(1), pts1X(2), pts1Y(2), pts2X(1), pts2Y(1), pts2X(2), pts2Y(2))
  pts(2) = Intercept(pts1X(4), pts1Y(4), pts1X(3), pts1Y(3), pts2X(1), pts2Y(1), pts2X(2), pts2Y(2))
  pts(3) = Intercept(pts1X(4), pts1Y(4), pts1X(3), pts1Y(3), pts2X(4), pts2Y(4), pts2X(3), pts2Y(3))
  pts(4) = Intercept(pts1X(1), pts1Y(1), pts1X(2), pts1Y(2), pts2X(4), pts2Y(4), pts2X(3), pts2Y(3))
  'desenhar
  Desenho.PicBox.FillStyle = 0
  Desenho.PicBox.FillColor = 0
  Desenho.PicBox.ForeColor = 0
  Call Polygon(Desenho.PicBox.hDC, pts(1), 4)
  Desenho.PicBox.DrawWidth = 1
  Desenho.PicBox.ForeColor = 0
End Sub

Public Sub DrawAllJunctions()
  Dim actP As Parede, nxtP As Parede
  Dim i As Long, j As Long
  
  '***opção junção
  If Paredes.Count = 0 Then Exit Sub
  'para todas as paredes, fazer junção nos dois pontos
  For i = 1 To Paredes.Count
    Set actP = Paredes.Item(i)
    For j = i + 1 To Paredes.Count
       Set nxtP = Paredes.Item(j)
       If ((actP.X1 = nxtP.X1) And (actP.Y1 = nxtP.Y1)) Or _
          ((actP.X1 = nxtP.X2) And (actP.Y1 = nxtP.Y2)) Or _
          ((actP.X2 = nxtP.X1) And (actP.Y2 = nxtP.Y1)) Or _
          ((actP.X2 = nxtP.X2) And (actP.Y2 = nxtP.Y2)) _
       Then Call Junction(actP, nxtP)
    Next j
  Next i
End Sub

'+--------------------------------------------------------+
'| Tratamento de Paredes                                  |
'+--------------------------------------------------------+
Public Sub EliminarParede(ByVal Num As Long)
  Dim i As Long
  
  Paredes.Remove Num
  SelectP = 0
  For i = Janelas.Count To 1 Step -1
    If Janelas.Item(i).ParedeNum = Num Then Janelas.Remove Num
  Next i
End Sub

Public Sub DrawAllParedes()
  Dim i As Integer
  
  'desenhar todas as paredes
  If Paredes.Count > 0 Then
    For i = 1 To Paredes.Count
      If (Design.TBTool = 1) Or (Design.TBTool = 3) Then
        Paredes.Item(i).DrawParede Desenho, False, 0, False
      Else
        Paredes.Item(i).DrawParede Desenho, False, RGB(240, 240, 240), False
      End If
    Next i
  End If
  'desenhar junções de paredes
  Call DrawAllJunctions
  'desenhar parede seleccionada
  If SelectP <> 0 Then Paredes.Item(SelectP).DrawParede Desenho, True, 0, False
End Sub

Public Sub GetPointParede(ByVal X As Long, ByVal Y As Long)
  Dim i As Integer
  Dim dist As Long
  Dim DistMin As Long
    
  SelectP = 0
  DistMin = 9999999
  If Paredes.Count > 0 Then
    For i = 1 To Paredes.Count
      dist = Paredes.Item(i).DistanciaPonto(X, Y)
      If dist <= DistMin Then
        SelectP = i
        DistMin = dist
      End If
    Next i
    If SelectP > 0 Then
      If DistMin > (Paredes.Item(SelectP).Largura \ 2) Then SelectP = 0
    End If
  End If
End Sub


'+--------------------------------------------------------+
'| Tratamento de Polígonos                                |
'+--------------------------------------------------------+
Public Sub DrawPolygon(ByRef PolP() As Point, mPo As Integer, Cf As Long, Cd As Long, Dp As Byte)
  Dim Bool As Long
  Dim i As Integer

  If mPo > 2 Then
    Design.Plan.FillStyle = 0
    Design.Plan.FillColor = Cf
    Design.Plan.ForeColor = Cf
    Bool = Polygon(Design.Plan.hDC, PolP(1), mPo)
    Design.Plan.FillStyle = Dp
    Design.Plan.FillColor = Cd
    Design.Plan.ForeColor = Cf
    Bool = Polygon(Design.Plan.hDC, PolP(1), mPo)
  End If
End Sub

Public Sub DrawAllPolygon()
  Dim PolP(200) As Point
  Dim i As Integer
  Dim j As Integer
  
  For j = 1 To Polygons
    For i = 1 To PolyMax(j)
      PolP(i).X = Desenho.MapToScreen(PolyG(j, i).X, 1)
      PolP(i).Y = Desenho.MapToScreen(PolyG(j, i).Y, 2)
    Next i
    Call DrawPolygon(PolP, PolyMax(j), PolyCF(j), PolyCD(j), PolyD(j))
  Next j
End Sub

Public Sub DrawSemiPolygon()
  Dim i As Integer
  Dim X1 As Integer
  Dim Y1 As Integer
  Dim X2 As Integer
  Dim Y2 As Integer
  
  For i = 1 To MaxPoints - 1
     X1 = Desenho.MapToScreen(PolyP(i).X, 1)
     Y1 = Desenho.MapToScreen(PolyP(i).Y, 2)
     X2 = Desenho.MapToScreen(PolyP(i + 1).X, 1)
     Y2 = Desenho.MapToScreen(PolyP(i + 1).Y, 2)
     Design.Plan.ForeColor = RGB(0, 250, 100)
     Design.Plan.Line (X1 - 1, Y1 - 1)-(X1 + 1, Y1 + 1), 0, B
     Design.Plan.Line (X1, Y1)-(X2, Y2)
  Next i
  Design.Plan.Line (X2 - 1, Y2 - 1)-(X2 + 1, Y2 + 1), 0, B
End Sub

Public Function LastPolyPoint() As Boolean
  Dim i As Integer
  
  If (PolyP(MaxPoints).X = PolyP(1).X) And (PolyP(MaxPoints).Y = PolyP(1).Y) Then
    LastPolyPoint = True
    Polygons = Polygons + 1
    Design.Gravar = True
    Call SetGravar
    PolyMax(Polygons) = MaxPoints - 1
    For i = 1 To MaxPoints - 1
      PolyG(Polygons, i).X = PolyP(i).X
      PolyG(Polygons, i).Y = PolyP(i).Y
    Next i
    PolyCF(Polygons) = CorF
    PolyCD(Polygons) = CorD
    PolyD(Polygons) = Dbox(DesP)
    MaxPoints = 0
  Else
    LastPolyPoint = False
  End If
End Function

'+--------------------------------------------------------+
'| Tratamento de Janelas                                  |
'+--------------------------------------------------------+
Public Sub TrocaPositionJanelas(ByVal Num As Long)
  Dim i As Long
  Dim Tam As Long
  
  Tam = Paredes.Item(Num).Tamanho
  For i = Janelas.Count To 1 Step -1
    If (Janelas.Item(i).ParedeNum = Num) Then
      Janelas.Item(i).Position = Tam - Janelas.Item(i).Position - Janelas.Item(i).Tamanho
    End If
  Next i
End Sub


Public Sub ReposicionarJanelas(ByVal Num As Long)
  Dim i As Long
  Dim Tam As Long
  
  Tam = Paredes.Item(Num).Tamanho
  For i = Janelas.Count To 1 Step -1
    If (Janelas.Item(i).ParedeNum = Num) And (Janelas.Item(i).Position + Janelas.Item(i).Tamanho > Tam) Then
      If Janelas.Item(i).Tamanho > Tam Then
        SelectJ = 0
        Janelas.Remove i
      Else
        Janelas.Item(i).Position = Tam - Janelas.Item(i).Tamanho
      End If
    End If
  Next i
End Sub

Public Sub GetPointJanela(ByVal X As Long, ByVal Y As Long)
  Dim i As Integer
  Dim dist As Long
  Dim DistMin As Long
  Dim Larg As Long
    
  SelectJ = 0
  DistMin = 9999999
  If Janelas.Count > 0 Then
   For i = 1 To Janelas.Count
     Larg = Janelas.Item(i).Suporte.Largura
     dist = Janelas.Item(i).DistanciaPonto(X, Y)
     If (dist <= DistMin) And (dist <= Larg) Then
       SelectJ = i
       DistMin = dist
     End If
   Next i
  End If
End Sub

Public Sub InsertWindow(Num As Integer)
  Dim Res As Boolean
  Dim NovaJanela As New Janela
  Dim TamRef As Long
  
  TamRef = TJanela.JWidth
  If Paredes.Item(Num).Tamanho < TamRef Then
    Call ErroTool(901, 0)
  Else
    Set NovaJanela.Suporte = Paredes.Item(Num)
    NovaJanela.ParedeNum = Num
    NovaJanela.Position = 0
    NovaJanela.Tamanho = TJanela.JWidth
    Janelas.Add NovaJanela
    Set NovaJanela = Nothing
  End If
End Sub

Public Sub DrawAllJanelas()
  Dim i As Integer
  
  If Janelas.Count > 0 Then
    For i = 1 To Janelas.Count
      Janelas.Item(i).DrawJanela Desenho, (i = SelectJ)
    Next i
  End If
End Sub

'+--------------------------------------------------------+
'| Desenhar a Planta                                      |
'+--------------------------------------------------------+
Public Sub DrawDesignPlan()
  Design.Plan.Cls
  Design.RulersH.Draw Desenho.Zoom, Desenho.MeioX, Desenho.CentroX, True, Desenho.TamanhoY, Design.Plan
  Design.RulersV.Draw Desenho.Zoom, Desenho.MeioY, Desenho.CentroY, True, Desenho.TamanhoX, Design.Plan
  Select Case Design.TBTool
    Case 1
      DrawAllPolygon
      DrawAllParedes
      DrawAllJanelas
    Case 2
      DrawAllParedes
      DrawAllPolygon
    Case 3
      DrawAllPolygon
      DrawAllParedes
      DrawAllJanelas
    Case 4
      DrawAllParedes
      DrawAllPolygon
  End Select
End Sub
