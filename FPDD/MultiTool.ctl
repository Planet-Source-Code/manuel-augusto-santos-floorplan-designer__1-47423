VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MultiTool 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2040
   ScaleHeight     =   141
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   136
   ToolboxBitmap   =   "MultiTool.ctx":0000
   Begin MSComctlLib.ImageList ToolListPic 
      Left            =   -15000
      Top             =   1215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Dummy 
      Height          =   330
      Left            =   -15000
      TabIndex        =   4
      Top             =   1800
      Width           =   75
   End
   Begin VB.PictureBox ToolsPic 
      AutoRedraw      =   -1  'True
      Height          =   1860
      Left            =   0
      ScaleHeight     =   120
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   225
      Width           =   1995
      Begin MSComctlLib.Toolbar SubPic 
         Height          =   1710
         Left            =   0
         TabIndex        =   5
         Top             =   -30
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   3016
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         ImageList       =   "ToolListPic"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton SubDesc 
         Height          =   330
         Index           =   0
         Left            =   345
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   1350
      End
      Begin VB.VScrollBar VScroll 
         Height          =   1680
         Left            =   1710
         Max             =   1
         Min             =   1
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Value           =   1
         Width           =   210
      End
   End
   Begin VB.CommandButton MainTool 
      Height          =   240
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1995
   End
End
Attribute VB_Name = "MultiTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const MainHeightDefault = 16
Const SelectColorDefault = &HD6EA97
Const MaxToolsDefault = 0
'Private properties
Private Started As Boolean
Private MainPos As Integer
Private MaxButtons As Byte
Private MaxTools(0 To 255) As Byte
Private ControlLock As Boolean
'Event Declarations:
Event ToolClick(ByVal MainSection As Byte, ByVal Button As Byte)
Event PropertiesClick(ByVal MainSection As Byte, ByVal Button As Byte)
'Property Variables:
Private mMainHeight As Long
Private mSelectColor As OLE_COLOR

'----------------------------------------------- PUBLIC
Public Sub DrawTools(ByVal modo As Byte)
  'modo: 1-full draw 2-redraw positions
  Dim i As Integer
  Dim Num As Integer
  Dim PicAux As Picture
 
  VScroll.Visible = False
  'limpa buttons
  If modo = 1 Then
    SubPic.Buttons.Clear
    If SubDesc.Count > 1 Then
      For i = SubDesc.UBound To 1 Step -1
        Unload SubDesc(i)
      Next i
    End If
    SubDesc(0).Visible = False
  End If
  'posiciona controlos antes da toolbox
  For i = 0 To MainPos
    MainTool(i).top = i * MainTool(0).Height
    If modo = 1 Then MainTool(i).Caption = LoadResString(10001 + i)
  Next i
  'posiciona toolbox
  ToolsPic.top = (MainPos + 1) * MainTool(0).Height
  i = UserControl.ScaleHeight - MaxButtons * MainTool(0).Height
  If i > 2 Then ToolsPic.Height = i Else ToolsPic.Height = 2
  VScroll.Height = ToolsPic.ScaleHeight
  'posiciona controlos depois da toolbox
  If (MainPos < MaxButtons - 1) Then
    For i = MainPos + 1 To MaxButtons - 1
      MainTool(i).top = ToolsPic.top + ToolsPic.Height + (i - MainPos - 1) * MainTool(0).Height
      If modo = 1 Then MainTool(i).Caption = LoadResString(10001 + i)
    Next i
  End If
  'determina estado da scroll bar
  VScroll.Enabled = False
  If MaxTools(MainPos) > 0 Then
    If MaxTools(MainPos) * 22 - ToolsPic.ScaleHeight > 0 Then
      VScroll.Enabled = True
      VScroll.Max = MaxTools(MainPos)
    End If
  End If
  'coloca icons das ferramentas na lista
  If modo = 1 Then
    SubPic.ImageList = Nothing
    ToolListPic.ListImages.Clear
    ToolListPic.ImageHeight = 16
    ToolListPic.ImageWidth = 16
    For i = 1 To MaxTools(MainPos)
      ToolListPic.ListImages.Add i, , LoadResPicture((MainPos + 1) * 100 + i, vbResIcon)
    Next i
  End If
  'cria buttons para toolbox
  If MaxTools(MainPos) > 0 Then
    If modo = 1 Then SubPic.ImageList = ToolListPic
    For i = 1 To MaxTools(MainPos)
      Num = (MainPos + 1) * 100 + i + 1000
      If i = 1 Then SubDesc(0).Caption = LoadResString(Num)
      If i > 1 Then
        If modo = 1 Then
          Load SubDesc(i - 1)
          SubDesc(i - 1).Caption = LoadResString(Num)
        End If
        SubDesc(i - 1).Left = SubDesc(0).Left
        SubDesc(i - 1).Width = SubDesc(0).Width
      End If
      SubDesc(i - 1).Visible = True
      SubDesc(i - 1).top = (i - VScroll.Value) * 22
      If modo = 1 Then SubPic.Buttons.Add , , , , i
    Next i
  End If
  SubPic.top = -2 - (VScroll.Value - 1) * 22
  'Põe cor diferente na ferramenta activa
  For i = 0 To MainTool.UBound
    MainTool(i).BackColor = vbButtonFace
  Next i
  MainTool(MainPos).BackColor = mSelectColor
  'ajustes
  VScroll.Visible = True
  If Started Then Dummy.SetFocus
End Sub

Public Sub SetMaxButtons(ByVal Max As Byte)
  Dim MaxIndex As Integer
  Dim i As Integer
  
  ControlLock = True
  MaxButtons = Max
  If MaxButtons = 0 Then MaxButtons = 1
  'limpar butões principais
  If MainTool.Count > 1 Then
    For i = MainTool.UBound To 1 Step -1
      Unload MainTool(i)
    Next i
  End If
  'carregar o controlo com butões
  For i = 1 To MaxButtons - 1
    Load MainTool(i)
    MainTool(i).Left = MainTool(0).Left
    MainTool(i).Width = MainTool(0).Width
    MainTool(i).Visible = True
  Next i
  ControlLock = False
End Sub

Public Sub SetMaxTools(ByRef numTools() As Byte)
  Dim i As Byte
  
  For i = 1 To MaxButtons
    MaxTools(i - 1) = numTools(i - 1)
  Next i
End Sub

Public Sub Start()
  Call DrawTools(1)
  Started = True
End Sub

'----------------------------------------------- PRIVATE
Private Sub MainTool_Click(Index As Integer)
  ControlLock = True
  MainPos = Index
  VScroll = 1
  Call DrawTools(1)
  ControlLock = False
End Sub

Private Sub MainTool_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Started Then Dummy.SetFocus
End Sub

Private Sub SubDesc_Click(Index As Integer)
  If Started Then Dummy.SetFocus
  RaiseEvent ToolClick(MainPos + 1, Index + 1)
End Sub

Private Sub SubDesc_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Started Then Dummy.SetFocus
End Sub

Private Sub SubPic_ButtonClick(ByVal Button As MSComctlLib.Button)
  RaiseEvent PropertiesClick(MainPos + 1, Button.Index)
End Sub

Private Sub SubPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Started Then Dummy.SetFocus
End Sub

Private Sub ToolsPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Started Then Dummy.SetFocus
End Sub

Private Sub UserControl_Initialize()
  MainPos = 0
  MaxButtons = 0
  Started = False
  ControlLock = False
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Started Then Dummy.SetFocus
End Sub

Private Sub UserControl_Resize()
  ControlLock = True
  VScroll = 1
  MainTool(0).Width = UserControl.ScaleWidth
  ToolsPic.Width = UserControl.ScaleWidth
  VScroll.Left = ToolsPic.ScaleWidth - VScroll.Width
  SubDesc(0).Width = VScroll.Left - SubDesc(0).Left
  ToolsPic.Height = UserControl.ScaleHeight - ToolsPic.top
  VScroll.Height = ToolsPic.ScaleHeight - VScroll.top
  If Started = True Then Call DrawTools(2)
  ControlLock = False
End Sub

Private Sub VScroll_Change()
  If ControlLock Then Exit Sub
  Call DrawTools(2)
End Sub

'----------------------------------------------- PROPERTIES
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get SelectColor() As OLE_COLOR
    SelectColor = mSelectColor
End Property

Public Property Let SelectColor(ByVal New_SelectColor As OLE_COLOR)
    mSelectColor = New_SelectColor
    PropertyChanged "SelectColor"
    MainTool(MainPos).BackColor = mSelectColor
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    mSelectColor = SelectColorDefault
    mMainHeight = MainHeightDefault
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    SelectColor = PropBag.ReadProperty("SelectColor", SelectColorDefault)
    MainHeight = PropBag.ReadProperty("MainHeight", MainHeightDefault)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("SelectColor", mSelectColor, SelectColorDefault)
    Call PropBag.WriteProperty("MainHeight", mMainHeight, MainHeightDefault)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,22
Public Property Get MainHeight() As Long
  MainHeight = mMainHeight
End Property

Public Property Let MainHeight(ByVal New_MainHeight As Long)
  mMainHeight = New_MainHeight
  PropertyChanged "MainHeight"
  MainTool(0).Height = mMainHeight
  ToolsPic.top = MainTool(0).Height - 1
  ToolsPic.Height = UserControl.ScaleHeight - ToolsPic.top
  VScroll.Height = ToolsPic.ScaleHeight - VScroll.top
  If Started = True Then Call DrawTools(2)
End Property

