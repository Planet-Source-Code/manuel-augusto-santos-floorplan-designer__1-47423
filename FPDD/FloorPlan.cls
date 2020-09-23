VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FloorPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Suporte para o desenho"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarTamanhoX As Long
Private mvarTamanhoY As Long
Private mvarCentroX As Long
Private mvarCentroY As Long
Private mvarMeioX As Long
Private mvarMeioY As Long
Private mvarZoom As Integer
Private ScaleMinPTS(11) As Integer
Private ScalePixCTM(11) As Single
Private ScaleMinCTM(11) As Integer
Public PicBox As PictureBox

'----------------------------------------------- PUBLIC
Public Sub SetDrawingIN(ByRef X As Long, ByRef Y As Long)
  If X < 0 Then X = 0
  If X > mvarTamanhoX Then X = mvarTamanhoX
  If Y < 0 Then Y = 0
  If Y > mvarTamanhoY Then Y = mvarTamanhoY
End Sub

Public Function DrawingIN(ByVal X As Long, ByVal Y As Long) As Boolean
  DrawingIN = False
  If (X >= 0) And (X <= mvarTamanhoX) And (Y >= 0) And (Y <= mvarTamanhoY) Then DrawingIN = True
End Function

Public Function MapToScreen(ByVal Pos As Single, ByVal modo As Byte) As Long
  Dim Valor As Single
  
  If modo = 1 Then
    Valor = Pos + mvarCentroX - mvarMeioX
    Valor = (Valor * ScaleMinPTS(mvarZoom)) / ScaleMinCTM(mvarZoom)
    Valor = Valor + mvarTamanhoX / 2
  Else
    Valor = Pos + mvarCentroY - mvarMeioY
    Valor = (Valor * ScaleMinPTS(mvarZoom)) / ScaleMinCTM(mvarZoom)
    Valor = Valor + mvarTamanhoY / 2
  End If
  MapToScreen = Round(Valor, 0)
End Function

'----------------------------------------------- SCALE VALUES
Public Sub SetScaleValue(ByRef MinPTS() As Integer, ByRef PixCTM() As Single, ByRef MinCTM() As Integer)
  Dim i As Byte
  
  For i = 0 To 10
     ScaleMinPTS(i) = MinPTS(i)
     ScalePixCTM(i) = PixCTM(i)
     ScaleMinCTM(i) = MinCTM(i)
  Next i
End Sub

Public Function MinPTS() As Integer
  MinPTS = ScaleMinPTS(mvarZoom)
End Function

Public Function PixCTM() As Single
  PixCTM = ScalePixCTM(mvarZoom)
End Function

Public Function MinCTM() As Integer
  MinCTM = ScaleMinCTM(mvarZoom)
End Function

'----------------------------------------------- ZOOM
Public Property Let Zoom(ByVal vData As Integer)
    mvarZoom = vData
End Property

Public Property Get Zoom() As Integer
    Zoom = mvarZoom
End Property

'----------------------------------------------- MEIO
Public Sub Meio(ByVal vDataX As Long, ByVal vDataY As Long)
    mvarMeioX = vDataX
    mvarMeioY = vDataY
End Sub

Public Property Get MeioY() As Long
    MeioY = mvarMeioY
End Property

Public Property Get MeioX() As Long
    MeioX = mvarMeioX
End Property

'----------------------------------------------- CENTRO
Public Sub Centro(ByVal vDataX As Long, ByVal vDataY As Long)
    mvarCentroX = vDataX
    mvarCentroY = vDataY
End Sub

Public Property Get CentroY() As Long
    CentroY = mvarCentroY
End Property

Public Property Get CentroX() As Long
    CentroX = mvarCentroX
End Property

'----------------------------------------------- TAMANHO
Public Sub Tamanho(ByVal vDataX As Long, ByVal vDataY As Long)
    mvarTamanhoX = vDataX
    mvarTamanhoY = vDataY
End Sub

Public Property Get TamanhoY() As Long
    TamanhoY = mvarTamanhoY
End Property

Public Property Get TamanhoX() As Long
    TamanhoX = mvarTamanhoX
End Property



