Attribute VB_Name = "MfileIO"
Option Explicit

Public Sub DesignLoadFile(fich As String)
  Dim Marca As String
  Dim F As Integer
  Dim i As Long
  Dim j As Long
  Dim NumElem As Long
  Dim ValorI As Integer
  Dim ValorL As Long
  Dim NovaParede As Parede
  Dim NovaJanela As Janela
  
  F = FreeFile()
  Open fich For Binary As F
  'obter a marca de controlo no ficheiro para ver se é válido
  Marca = "1234567890"
  Get F, 1, Marca
  If Marca <> "FPlan v1.0" Then
    '********* erro
    Exit Sub
  End If
  'ler os itens do ficheiro
  On Error GoTo FechaFich
LerItem:
  Get F, , Marca
  Select Case Marca
    'ler paredes
    Case "[PAREDES ]"
      Get F, , NumElem
      For i = 1 To NumElem
        Set NovaParede = New Parede
        Get F, , ValorI: NovaParede.Largura = ValorI
        Get F, , ValorL: NovaParede.X1 = ValorL
        Get F, , ValorL: NovaParede.Y1 = ValorL
        Get F, , ValorL: NovaParede.X2 = ValorL
        Get F, , ValorL: NovaParede.Y2 = ValorL
        Paredes.Add NovaParede
        Set NovaParede = Nothing
      Next i
      GoTo LerItem
    'ler polígonos
    Case "[POLYGONS]"
      Get F, , Polygons
      For i = 1 To Polygons
        Get F, , PolyCF(i)
        Get F, , PolyCD(i)
        Get F, , PolyD(i)
        Get F, , PolyMax(i)
        For j = 1 To PolyMax(i)
          Get F, , PolyG(i, j).X
          Get F, , PolyG(i, j).Y
        Next j
      Next i
      GoTo LerItem
    'ler Janelas
    Case "[JANELAS ]"
      Get F, , NumElem
      For i = 1 To NumElem
        Set NovaJanela = New Janela
        Get F, , ValorL: NovaJanela.ParedeNum = ValorL
        Set NovaJanela.Suporte = Paredes.Item(NovaJanela.ParedeNum)
        Get F, , ValorL: NovaJanela.Position = ValorL
        Get F, , ValorL: NovaJanela.Tamanho = ValorL
        Janelas.Add NovaJanela
        Set NovaJanela = Nothing
      Next i
  End Select
  'fim da leitura do ficheiro
FechaFich:
  Close F
End Sub

Public Sub DesignSaveFile(fich As String)
  Dim Marca As String
  Dim i As Integer
  Dim j As Integer
  Dim F As Integer
  Dim ValorL As Long
  Dim ValorI As Integer
  
  Marca = "FPlan v1.0"
  F = FreeFile()
  Open fich For Binary As F
  'escrever a marca de controlo no ficheiro
  Put F, 1, Marca
  'escrever as paredes
  If Paredes.Count > 0 Then
    Marca = "[PAREDES ]"
    Put F, , Marca
    ValorL = Paredes.Count
    Put F, , ValorL
    For i = 1 To Paredes.Count
      ValorI = Paredes.Item(i).Largura
      Put F, , ValorI
      ValorL = Paredes.Item(i).X1
      Put F, , ValorL
      ValorL = Paredes.Item(i).Y1
      Put F, , ValorL
      ValorL = Paredes.Item(i).X2
      Put F, , ValorL
      ValorL = Paredes.Item(i).Y2
      Put F, , ValorL
    Next i
  End If
  'escrever os polígonos
  If Polygons > 0 Then
    Marca = "[POLYGONS]"
    Put F, , Marca
    Put F, , Polygons
    For i = 1 To Polygons
      Put F, , PolyCF(i)
      Put F, , PolyCD(i)
      Put F, , PolyD(i)
      Put F, , PolyMax(i)
      For j = 1 To PolyMax(i)
        Put F, , PolyG(i, j).X
        Put F, , PolyG(i, j).Y
      Next j
    Next i
  End If
  'escrever as Janelas
  If Janelas.Count > 0 Then
    Marca = "[JANELAS ]"
    Put F, , Marca
    ValorL = Janelas.Count
    Put F, , ValorL
    For i = 1 To Janelas.Count
      ValorL = Janelas.Item(i).ParedeNum
      Put F, , ValorL
      ValorL = Janelas.Item(i).Position
      Put F, , ValorL
      ValorL = Janelas.Item(i).Tamanho
      Put F, , ValorL
    Next i
  End If
  'fim da escrita do ficheiro
  Close F
End Sub
