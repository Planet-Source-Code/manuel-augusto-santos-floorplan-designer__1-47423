Attribute VB_Name = "Mfpdd"
'-------------------------------------------------------------------'
' Floor Plan Digital Designer
' Started 10/03/1999
' Manuel Augusto Nogueira dos Santos
' Mfpdd - Módulo para Design
'-------------------------------------------------------------------'
Option Explicit
'---------------------------------- Descrição de valores
' Oper :
'      0 - normal
'      ----------------
'      2 - Draw Paredes
'      ----------------
'      ----------------
'      5 - Drag Janela
'----------------------------------
'----------------------------------Estruturas
Public Type Rect
  Left As Long
  top As Long
  right As Long
  bottom As Long
End Type

Public Type Point
  X As Long
  Y As Long
End Type

'----------------------------------Variáveis Globais
Public PolyP(50) As Point ' posição dos pontos do polígono actual
Public CorF As Long ' cor de fundo dos polígonos
Public CorD As Long ' cor do desenho dos polígonos
Public DesP As Byte ' nº activo de desenho dos polígonos
Public MaxPoints As Long 'número de pontos do polígono
Public oldX As Long '\ antiga posição
Public oldY As Long '/ do mouse
Public Oper As Byte 'operação em curso
Public PIXEL(11) As Single 'centímetros por pixel
Public CENTIM(11) As Integer 'medida mínima em centímetros
Public PONTOS(11) As Integer 'número mínimo de pontos por unidade
Public GLock As Boolean 'fecho centralizado
Public MaxBT(10) As Byte 'número de items em cada graupo de ferramentas
Public Const MaxB As Integer = 10000 'Máximo em cm da área de desenho
Public Const DefDimWall As Byte = 20
Private LockOBJPos As Byte
'---------------------------------------------------

Public Sub Main()
  Dim i As Integer
    Design.Ready = False
    Splash.Show
    Splash.Refresh
    '------------------carregamento do programa
    GLock = False
    
    'Dados fixos para escalas de zoom
    PIXEL(0) = 0.1:    PIXEL(1) = 0.2:    PIXEL(2) = 0.4
    PIXEL(3) = 0.5:    PIXEL(4) = 1:      PIXEL(5) = 2
    PIXEL(6) = 4:      PIXEL(7) = 5:      PIXEL(8) = 10
    PIXEL(9) = 20:     PIXEL(10) = 40
    
    PONTOS(0) = 10:    PONTOS(1) = 5:     PONTOS(2) = 5
    PONTOS(3) = 10:    PONTOS(4) = 5:     PONTOS(5) = 5
    PONTOS(6) = 5:     PONTOS(7) = 5:     PONTOS(8) = 5
    PONTOS(9) = 5:     PONTOS(10) = 5:
    
    CENTIM(0) = 1:     CENTIM(1) = 1:     CENTIM(2) = 2
    CENTIM(3) = 5:     CENTIM(4) = 5:     CENTIM(5) = 10
    CENTIM(6) = 20:    CENTIM(7) = 25:    CENTIM(8) = 50
    CENTIM(9) = 100:   CENTIM(10) = 200
    
    'prepara área de desenho
    Desenho.SetScaleValue PONTOS, PIXEL, CENTIM
    Set Desenho.PicBox = Design.Plan
    Desenho.Zoom = 7
    Desenho.Meio MaxB \ 2, MaxB \ 2
    Desenho.Centro MaxB \ 2, MaxB \ 2
    Desenho.Tamanho 100, 100
    'prepara Rulers
    Design.RulersH.SetScale PONTOS, CENTIM
    Design.RulersV.SetScale PONTOS, CENTIM
    Design.RulersH.Active = True
    Design.RulersV.Active = True
    'prepara ToolBox
    MaxBT(0) = 6
    MaxBT(1) = 3
    MaxBT(2) = 2
    MaxBT(3) = 2
    MaxBT(6) = 3
    Design.DrawTool.SetMaxButtons 7
    Design.DrawTool.SetMaxTools MaxBT
    Design.DrawTool.Start
    
    SelectP = 0
    SelectJ = 0
    MaxPoints = 0
    'Paredes = 0
    Polygons = 0
    'Janelas2 = 0
    CorF = RGB(190, 250, 205)
    CorD = RGB(20, 55, 55)
    Dbox(0) = 0
    'ParedesL(0) = DefDimWall
    TJanela.JWidth = 150
    TJanela.JLeft = 5
    For i = 2 To 7: Dbox(i - 1) = i: Next i
    For i = 8 To 12: Dbox(i - 1) = 0: Next i
    DesP = 5
    '------------------início do programa
    Load Design
    Unload Splash
    Design.Show
    DoEvents
    Design.Ready = True
End Sub

Public Sub SetGravar()
  If Design.Gravar Then
    Design.StatusBar1.Panels("Gravar").Picture = LoadResPicture(1, vbResIcon)
  Else
    Design.StatusBar1.Panels("Gravar").Picture = LoadResPicture(2, vbResIcon)
  End If
End Sub

Public Sub SetPatternIcons()
  
  Exit Sub
  '    Design.Patterns(0).FillStyle = 0
  '    Design.Patterns(0).FillColor = CorF
  '    Design.Patterns(0).ForeColor = CorF
  '    Design.Patterns(0).Line (0, 0)-Step(16, 16), , B
  '    Design.Patterns(2).FillStyle = 0
  '    Design.Patterns(2).FillColor = CorF
  '    Design.Patterns(2).ForeColor = CorF
  '    Design.Patterns(2).Line (0, 0)-Step(16, 16), , B
  '    Design.Patterns(1).FillStyle = 0
  '    Design.Patterns(1).FillColor = CorD
  '    Design.Patterns(1).ForeColor = CorD
  '    Design.Patterns(1).Line (0, 0)-Step(16, 16), , B
  '    Design.Patterns(2).FillStyle = Dbox(DesP)
  '    Design.Patterns(2).ForeColor = CorF
  '    Design.Patterns(2).FillColor = CorD
  '    Design.Patterns(2).Line (0, 0)-Step(16, 16), , B
  '
  '    Design.Patterns(0).top = -100
  '    Design.Patterns(1).top = -100
  '    Design.Patterns(2).top = -100
  '      Design.TButtons.Buttons.Add , , BStr, 0, 0
  '      Design.Patterns(i - 1).top = -Topo - 19 + i * 22
End Sub

Public Sub ErroTool(MSGID As Integer, MSGD As Integer)
  If MSGD > 0 Then
    MsgBox LoadResString(MSGID) & LoadResString(MSGD), vbExclamation
  Else
    MsgBox LoadResString(MSGID), vbExclamation
  End If
End Sub

Public Sub LoadResStrings(FRM As Form)
    On Error Resume Next
    Dim ctl As Control
    Dim obj As Object
    Dim fnt As Object
    Dim sCtlType As String
    Dim nVal As Integer

    'set the form's caption
    FRM.Caption = LoadResString(CInt(FRM.Tag))
    'set the font
    Set fnt = FRM.Font
    fnt.Name = LoadResString(20)
    fnt.Size = CInt(LoadResString(21))
    'set the controls' captions using the caption
    'property for menu items and the Tag property
    'for all other controls
    For Each ctl In FRM.Controls
        Set ctl.Font = fnt
        sCtlType = TypeName(ctl)
        If sCtlType = "Label" Then
            ctl.Caption = LoadResString(CInt(ctl.Tag))
        ElseIf sCtlType = "Menu" Then
            ctl.Caption = LoadResString(CInt(ctl.Caption))
        ElseIf sCtlType = "TabStrip" Then
            For Each obj In ctl.Tabs
                obj.Caption = LoadResString(CInt(obj.Tag))
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "Toolbar" Then
            For Each obj In ctl.Buttons
                obj.ToolTipText = LoadResString(CInt(obj.ToolTipText))
            Next
        ElseIf sCtlType = "ListView" Then
            For Each obj In ctl.ColumnHeaders
                obj.Text = LoadResString(CInt(obj.Tag))
            Next
        Else
            nVal = 0
            nVal = Val(ctl.Caption)
            If nVal > 0 Then ctl.Caption = LoadResString(CInt(nVal))
            nVal = 0
            nVal = Val(ctl.ToolTipText)
            If nVal > 0 Then ctl.ToolTipText = LoadResString(CInt(nVal))
        End If
    Next
End Sub

Public Sub SetObjectPos()
  Dim pX As Long
  Dim pY As Long
  
  DoEvents
  If LockOBJPos = False Then
    LockOBJPos = True
    If Design.WindowState <> 1 Then 'minimized
      oldX = -1000: oldY = -1000
      Desenho.Zoom = Design.Zoom
      Design.Plan.Width = Design.ScaleWidth - 126
      Design.Plan.Height = Design.ScaleHeight - 60
      pX = Design.Plan.ScaleWidth
      pY = Design.Plan.ScaleHeight
      If (pX And 1) = 1 Then pX = pX - 1
      If (pY And 1) = 1 Then pY = pY - 1
      Desenho.Tamanho pX, pY
      Design.RulersV.Height = Design.Plan.Height
      Design.RulersH.Width = Design.Plan.Width
      pX = Desenho.MeioX
      pY = Desenho.MeioY
      Call ToGrid(pX, pY)
      Desenho.Meio pX, pY
      Design.ScrMapa.top = Design.Tools.ScaleHeight - Design.ScrMapa.Height - 1
      Design.ScrMapa.DrawMapa Desenho.MeioX, Desenho.MeioY, Desenho.TamanhoX * Desenho.PixCTM, Desenho.TamanhoY * Desenho.PixCTM
      DrawDesignPlan
    End If
    DoEvents
    LockOBJPos = False
  End If
End Sub

Public Sub ToGrid(pX, pY)
  Dim Valor As Integer
  
  Valor = Design.Zoom
  If Valor > 2 Then
    If (pX Mod CENTIM(Valor)) <> 0 Then pX = ((pX \ CENTIM(Valor)) + 1) * CENTIM(Valor)
    If (pY Mod CENTIM(Valor)) <> 0 Then pY = ((pY \ CENTIM(Valor)) + 1) * CENTIM(Valor)
  End If
End Sub

Public Function GetCurPos(ByVal Pos As Long, ByVal Tam As Integer, ByVal PosEcr As Integer, ByVal Centro As Integer, ByVal Escala As Byte) As Integer
  Dim Valor As Integer
  
  Valor = Pos - Tam \ 2
  If Valor > 0 Then Valor = (Valor + PONTOS(Escala) \ 2) \ PONTOS(Escala)
  If Valor < 0 Then Valor = (Valor - PONTOS(Escala) \ 2) \ PONTOS(Escala)
  Valor = Valor * PONTOS(Escala)
  GetCurPos = PosEcr - Centro + Valor * PIXEL(Escala)
End Function
