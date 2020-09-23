Attribute VB_Name = "MMisc"
Option Explicit

'Public Function Hipotenusa(ByVal Cat1 As Long, ByVal Cat2 As Long) As Single
' calcula a hipotenusa pelo teorema de pitágoras
'  Hipotenusa = Sqr(Cat1 * Cat1 + Cat2 * Cat2)
'End Function

Public Function StrExp(ByVal n As Integer, ByVal stam As Byte, ByVal modo As Byte) As String
' vai buscar uma string e amplia-a até um número de caracteres
' modo : 1-Left; 2-Right; 3-Center
  Dim Res As String
  Dim insL As Boolean
  
  If modo <> 2 Then insL = False Else insL = True
  Res = LoadResString(n)
  Do While Len(Res) < stam
    If insL Then Res = " " & Res Else Res = Res & " "
    If modo = 3 Then insL = Not insL
  Loop
  StrExp = Res
End Function

'Public Function Regra3(ByVal N11 As Single, ByVal N12 As Single, ByVal N21 As Single, ByVal N22 As Single, ByVal modo As Byte) As Single
' faz uma regra de 3 simples
' modo : é o número do elemento que é para calcular
'  Dim Res As Single
'
'  Select Case modo
'    Case 1: Res = (N12 * N21) / N22
'    Case 2: Res = (N11 * N22) / N21
'    Case 3: Res = (N11 * N22) / N12
'    Case 4: Res = (N12 * N21) / N11
'  End Select
'  Regra3 = Res
'End Function

'Public Sub Rotate2D(ByVal PosX As Integer, ByVal PosY As Integer, _
'                    ByVal DirX As Integer, ByVal DirY As Integer, _
'                    ByVal Grau As Single, _
'                    ByRef RotX As Integer, ByRef RotY As Integer)
' faz uma rotação a partir de um centro de um ponto no plano
' segundo um determinado grau de rotação
' Pos  : Centro de rotação
' Dir  : Posição do ponto a partir do centro
' Grau : Angulo de rotação em radianos
' Rot  : Resultado da rotação do ponto em relação ao centro
'  Dim RX As Single
'  Dim RY As Single
'
'  RX = PosX + (DirX - PosX) * Cos(Grau) - (DirY - PosY) * Sin(Grau)
'  RY = PosY + (DirX - PosX) * Sin(Grau) + (DirY - PosY) * Cos(Grau)
'  RotX = Round(RX, 0)
'  RotY = Round(RY, 0)
'End Sub

