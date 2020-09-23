VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Design 
   BackColor       =   &H8000000C&
   Caption         =   "Floor Plan Digital Designer"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7140
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00808000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   3  'Windows Default
   Begin FPDD.Rulers RulersV 
      Height          =   4155
      Left            =   1605
      Top             =   645
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   7329
      Numeros         =   "Main.frx":058A
      Orientation     =   1
      MaxValue        =   10000
   End
   Begin FPDD.Rulers RulersH 
      Height          =   285
      Left            =   1890
      Top             =   360
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   503
      Numeros         =   "Main.frx":0B6C
      MaxValue        =   10000
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "1019"
            Object.Tag             =   "1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "1020"
            Object.Tag             =   "2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "1021"
            Object.Tag             =   "3"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "1022"
            Object.Tag             =   "4"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "1023"
            Object.Tag             =   "6"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Undo"
            Object.ToolTipText     =   "1024"
            Object.Tag             =   "7"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Redo"
            Object.ToolTipText     =   "1025"
            Object.Tag             =   "8"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Point"
            Object.ToolTipText     =   "1026"
            Object.Tag             =   "9"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Wall"
            Object.ToolTipText     =   "1027"
            Object.Tag             =   "10"
            ImageIndex      =   9
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Floor"
            Object.ToolTipText     =   "1028"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   301
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.ListBox Andar 
         Height          =   285
         IntegralHeight  =   0   'False
         ItemData        =   "Main.frx":10C2
         Left            =   4185
         List            =   "Main.frx":10DB
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   30
         Width           =   510
      End
      Begin VB.TextBox TxtY 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   5580
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   30
         Width           =   780
      End
      Begin VB.TextBox TxtX 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   4770
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   30
         Width           =   780
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4815
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7726
            MinWidth        =   2381
            Key             =   "Texto"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   688
            MinWidth        =   688
            Key             =   "Gravar"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "NUM"
            Key             =   "X"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "NUM"
            Key             =   "Y"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1005
            MinWidth        =   1005
            TextSave        =   "23:58"
            Key             =   "Horas"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   5175
      Top             =   -90
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4050
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   12648447
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4545
      Top             =   -240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":10FB
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":120F
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1323
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":163F
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1A93
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1BA7
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1CBB
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1DCF
            Key             =   "Point"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":20EB
            Key             =   "Wall"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":253F
            Key             =   "Florr"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Plan 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   4155
      Left            =   1890
      MouseIcon       =   "Main.frx":269B
      MousePointer    =   99  'Custom
      ScaleHeight     =   273
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   645
      Width           =   5235
   End
   Begin VB.PictureBox Tools 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      Height          =   4455
      Left            =   0
      ScaleHeight     =   293
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   102
      TabIndex        =   2
      Top             =   360
      Width           =   1590
      Begin FPDD.Position Zoom 
         Height          =   375
         Left            =   30
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   661
         BorderStyle     =   1
         Value           =   6
         MouseIcon       =   "Main.frx":29A5
         MousePointer    =   99
         Object.ToolTipText     =   "6"
      End
      Begin FPDD.MultiTool DrawTool 
         Height          =   2445
         Left            =   -30
         TabIndex        =   8
         Top             =   -30
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   4313
      End
      Begin FPDD.Mapa ScrMapa 
         Height          =   1500
         Left            =   30
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   2646
         BackPicture     =   "Main.frx":2B07
         Escala          =   100
         ScrPicture      =   "Main.frx":5669
         ScrBorder       =   0
         MouseIcon       =   "Main.frx":CBEB
         MousePointer    =   99
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileNew 
         Caption         =   "1001"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "1003"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "1004"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "1006"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "1007"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1008"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1009"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1010"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "1011"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "1012"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1013"
      Begin VB.Menu mnuViewGrid 
         Caption         =   "1014"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "1015"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1016"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1017"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1017"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1018"
      End
   End
End
Attribute VB_Name = "Design"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Sub GetCursorPos Lib "user32" (lpPoint As Point)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point)
'--------------------------------------
Private LockK32 As Boolean
'--------------------------------------
Public Ready As Boolean
Public TBTool As Byte
Public Gravar As Boolean
Public Ficheiro As String

Private Sub DrawTool_PropertiesClick(ByVal MainSection As Byte, ByVal Button As Byte)
  GLock = True
  DoEvents
  Select Case MainSection
    Case 1
      Select Case Button
        Case 1
          TParede.BkWidth = TParede.ComboWidth
          If SelectP <> 0 Then TParede.ComboWidth = Paredes.Item(SelectP).Largura
          TParede.Show vbModal, Me
          If SelectP <> 0 Then Paredes.Item(SelectP).Largura = TParede.ComboWidth
          DrawDesignPlan
        Case 4
          TJanela.BkWidth = TJanela.JWidth
          TJanela.BkLeft = TJanela.JLeft
          If SelectJ <> 0 Then
            TJanela.JLeft = Janelas.Item(SelectJ).Position
            TJanela.JWidth = Janelas.Item(SelectJ).Tamanho
          End If
          TJanela.Show vbModal, Me
          If SelectJ <> 0 Then
            Janelas.Item(SelectJ).Position = TJanela.JLeft
            Janelas.Item(SelectJ).Tamanho = TJanela.JWidth
          End If
          DrawDesignPlan
      End Select
    Case 7
      Select Case Button
        Case 1
          dlgCommonDialog.Color = CorF
          dlgCommonDialog.Flags = cdlCCRGBInit
          dlgCommonDialog.ShowColor
          CorF = dlgCommonDialog.Color
          '******************
        Case 2
          dlgCommonDialog.Color = CorD
          dlgCommonDialog.Flags = cdlCCRGBInit
          dlgCommonDialog.ShowColor
          CorD = dlgCommonDialog.Color
          '******************
        Case 3
          Pattern.Show vbModal, Me
          '******************
      End Select
  End Select
  GLock = False
End Sub

Private Sub DrawTool_ToolClick(ByVal MainSection As Byte, ByVal Button As Byte)
  GLock = True
  DoEvents
  Select Case MainSection
    Case 1
      Select Case Button
        Case 4
          If SelectP = 0 Then
            Call ErroTool(900, 1004 + MainSection * 100)
          Else
            Call InsertWindow(SelectP)
            DrawDesignPlan
          End If
      End Select
  End Select
  GLock = False
End Sub

Private Sub Form_Load()
  Dim i As Byte
  
  GLock = True
  LoadResStrings Me
  LockK32 = False
  TBTool = 1
  SetObjectPos
  Gravar = False
  Call SetGravar
  Ficheiro = ""
  GLock = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Oper = 0
End Sub

Private Sub Form_Resize()
  Dim CheckResize As Boolean
  Dim H As Long
  Dim W As Long
  
  GLock = True
  CheckResize = False
  If WindowState = 0 Then
    With Design
      If Height < 5700 Then CheckResize = True
      If Width < 8500 Then CheckResize = True
    End With
  End If
  If CheckResize = True Then
    Height = 5700
    Width = 8500
    DoEvents
  End If
  SetObjectPos
  GLock = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Integer
  'close all sub forms
  For i = Forms.Count - 1 To 1 Step -1
      Unload Forms(i)
  Next
  End
End Sub

Private Sub mnuViewGrid_Click()
  If mnuViewGrid.Checked = False Then mnuViewGrid.Checked = True Else mnuViewGrid.Checked = False
  DrawDesignPlan
End Sub

Private Sub RulersH_ChangePos(NovoMeio As Long)
  GLock = True
  Desenho.Meio NovoMeio, Desenho.MeioY
  Call SetObjectPos
  GLock = False
End Sub

Private Sub RulersV_ChangePos(NovoMeio As Long)
  GLock = True
  Desenho.Meio Desenho.MeioX, NovoMeio
  Call SetObjectPos
  GLock = False
End Sub

Private Sub ScrMapa_ChangePos(ByVal X As Long, ByVal Y As Long)
  GLock = True
  Desenho.Meio X, Y
  SetObjectPos
  GLock = False
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim i As Byte
  On Error Resume Next
  Select Case Button.Key
    Case "New":    mnuFileNew_Click
    Case "Open":   mnuFileOpen_Click
    Case "Save":   mnuFileSave_Click
    Case "Print":  mnuFilePrint_Click
    Case "Delete": 'mnuEditDelete_Click
    Case "Floor"
      TBTool = 2
      SelectP = 0
      SelectJ = 0
      DrawDesignPlan
    Case "Wall"
      TBTool = 1
      SelectP = 0
      SelectJ = 0
      DrawDesignPlan
    Case "Point"
      If TBTool = 1 Then
        TBTool = 3
        DrawDesignPlan
      Else
        TBTool = 4
        DrawDesignPlan
      End If
  End Select
End Sub

Private Sub mnuHelpAbout_Click()
    About.Show vbModal, Me
End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer
    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuViewOptions_Click()
    Options.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub

Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    'ToDo: Add 'mnuFilePrint_Click' code.
    MsgBox "Add 'mnuFilePrint_Click' code."
End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
  GLock = True
  dlgCommonDialog.DialogTitle = "Save As"
  dlgCommonDialog.CancelError = False
  dlgCommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames
  dlgCommonDialog.Filter = "Floor Plan Design (*.FPD)|*.FPD"
  dlgCommonDialog.ShowSave
  If Len(dlgCommonDialog.FileName) = 0 Then Exit Sub
  Ficheiro = dlgCommonDialog.FileName
  StatusBar1.Panels(1) = Ficheiro
  Call DesignSaveFile(Ficheiro)
  Design.Gravar = False
  Call SetGravar
  GLock = False
End Sub

Private Sub mnuFileSave_Click()
  GLock = True
  If Ficheiro <> "" Then
    Call DesignSaveFile(Ficheiro)
    Design.Gravar = False
    Call SetGravar
  Else
    mnuFileSaveAs_Click
    Exit Sub
  End If
  Design.Gravar = False
  Call SetGravar
  GLock = False
End Sub

Private Sub mnuFileOpen_Click()
  Dim sFile As String
  Dim i As Long

  dlgCommonDialog.DialogTitle = "Open"
  dlgCommonDialog.CancelError = False
  dlgCommonDialog.Flags = cdlOFNHideReadOnly Or cdlOFNLongNames
  dlgCommonDialog.Filter = "Floor Plan Design (*.FPD)|*.FPD"
  dlgCommonDialog.ShowOpen
  If Len(dlgCommonDialog.FileName) = 0 Then Exit Sub
  sFile = dlgCommonDialog.FileName
  GLock = True
  
  Desenho.Meio MaxB \ 2, MaxB \ 2
  Desenho.Centro MaxB \ 2, MaxB \ 2
  'limpa paredes
  Do While Paredes.Count > 0
    Paredes.Remove 1
  Loop
  'limpa janelas
  Do While Janelas.Count > 0
    Janelas.Remove 1
  Loop
  
  SelectP = 0
  SelectJ = 0
  Polygons = 0
  MaxPoints = 0
  Call DesignLoadFile(sFile)
  DrawDesignPlan
  Ficheiro = sFile
  StatusBar1.Panels(1) = Ficheiro
  Design.Refresh
  Design.Gravar = False
  Call SetGravar
  GLock = False
End Sub

Private Sub mnuFileNew_Click()
  Dim i As Integer
   
  GLock = True
  
  Desenho.Meio MaxB \ 2, MaxB \ 2
  Desenho.Centro MaxB \ 2, MaxB \ 2
  'limpa paredes
  Do While Paredes.Count > 0
    Paredes.Remove 1
  Loop
  'limpa janelas
  Do While Janelas.Count > 0
    Janelas.Remove 1
  Loop
  SelectP = 0
  SelectJ = 0
  MaxPoints = 0
  Polygons = 0
  Ficheiro = ""
  StatusBar1.Panels(1) = Ficheiro
  Design.Gravar = False
  Call SetGravar
  DrawDesignPlan
  GLock = False
End Sub

Private Sub Plan_KeyDown(KeyCode As Integer, Shift As Integer)
  GLock = True
  If (Zoom > 1) And (KeyCode = 32) And (LockK32 = False) And (Oper = 2) Then
    LockK32 = True
    Zoom = Zoom - 1
    DoEvents
    'reajustar posição
    DrawDesignPlan
  End If
  GLock = False
End Sub

Private Sub Plan_KeyUp(KeyCode As Integer, Shift As Integer)
  GLock = True
  If (Zoom < 10) And (KeyCode = 32) And (LockK32 = True) And (Oper = 2) Then
    LockK32 = False
    Zoom = Zoom + 1
    DoEvents
    'reajustar posição
    DrawDesignPlan
  End If
  GLock = False
End Sub

Private Sub Plan_MouseDown(Button As Integer, Shift As Integer, Xx As Single, Yy As Single)
  Dim Wpos As Point
  Dim Mpos As Point
  Dim X As Long
  Dim Y As Long
  Dim Xaux As Long
  Dim Yaux As Long
  Dim NovaParede As Parede
  
  GLock = True
  Select Case Oper
    Case 0
      Call GetCursorPos(Mpos)
      Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
      Call ClientToScreen(Plan.hWnd, Wpos)
      'determinar posição em cm
      X = GetCurPos(Mpos.X - Wpos.X, Desenho.TamanhoX, Desenho.MeioX, Desenho.CentroX, Zoom)
      Y = GetCurPos(Mpos.Y - Wpos.Y, Desenho.TamanhoY, Desenho.MeioY, Desenho.CentroY, Zoom)
      Xaux = X
      Yaux = Y
      Call ToGrid(X, Y) 'alinhar à grelha
      'determinar acção
      Select Case TBTool
        Case 1 'desenhar paredes (nova)
          Oper = 2
          Set NovaParede = New Parede
          NovaParede.Largura = TParede.ComboWidth
          NovaParede.X1 = X
          NovaParede.Y1 = Y
          NovaParede.X2 = X
          NovaParede.Y2 = Y
          Paredes.Add NovaParede
          Set NovaParede = Nothing
          SelectP = Paredes.Count
        Case 2 'desenhar o chão
          If MaxPoints <= 50 Then MaxPoints = MaxPoints + 1
          If MaxPoints = 51 Then MaxPoints = 50
          PolyP(MaxPoints).X = X
          PolyP(MaxPoints).Y = Y
          If MaxPoints > 1 Then
            If LastPolyPoint = True Then Call DrawDesignPlan Else Call DrawSemiPolygon
          End If
          Plan.Refresh
        Case 3 'Apontar aos objectos
          'Apontar à Janela
          If SelectJ > 0 Then
            If Janelas.Item(SelectJ).GotGrip(Xaux, Yaux) Then
              Oper = 5
              TJanela.JWidth = Janelas.Item(SelectJ).Tamanho
              TJanela.JLeft = Janelas.Item(SelectJ).Position
              GoTo fimpoint
            Else
              SelectJ = 0
            End If
          Else
            Call GetPointJanela(Xaux, Yaux)
            If SelectJ <> 0 Then
              SelectP = 0
              TJanela.JWidth = Janelas.Item(SelectJ).Tamanho
              TJanela.JLeft = Janelas.Item(SelectJ).Position
              GoTo fimpoint
            End If
          End If
          'Apontar à parede ------------------------------------------
          If SelectP > 0 Then
            Select Case Paredes.Item(SelectP).GotGrip(Xaux, Yaux)
              Case 0 'não clickou na parede
                SelectP = 0
              Case 1 'clickou em x1,y1
                Paredes.Item(SelectP).TrocaCoord
                Call TrocaPositionJanelas(SelectP)
                Oper = 2
                TParede.ComboWidth = Paredes.Item(SelectP).Largura
                Paredes.Item(SelectP).Visible = False
              Case 2 'clickou em x2,y2
                Oper = 2
                TParede.ComboWidth = Paredes.Item(SelectP).Largura
                Paredes.Item(SelectP).Visible = False
            End Select
          Else
            Call GetPointParede(Xaux, Yaux)
            If SelectP <> 0 Then
              TParede.ComboWidth = Paredes.Item(SelectP).Largura
              SelectJ = 0
            End If
          End If
        Case 4 'point chão
      End Select
    Case Else
      Oper = 0
  End Select
fimpoint:
  DrawDesignPlan
  If Oper = 2 Then Paredes.Item(SelectP).Visible = True
  GLock = False
End Sub

Private Sub Plan_MouseUp(Button As Integer, Shift As Integer, Xx As Single, Yy As Single)
  GLock = True
  If Oper = 2 Then
    Design.Gravar = True
    Call SetGravar
    If Paredes.Item(SelectP).Tamanho = 0 Then Call EliminarParede(SelectP)
    If SelectP > 0 Then Call ReposicionarJanelas(SelectP)
  End If
  If TBTool = 1 Then SelectP = 0
  DrawDesignPlan
  Oper = 0
  GLock = False
End Sub

Private Sub Timer1_Timer()
  Dim X As Long
  Dim Y As Long
  Dim GridX As Long
  Dim GridY As Long
  Dim Mpos As Point
  Dim Wpos As Point
  Dim DifX As Integer
  Dim DifY As Integer
    
  If GLock = True Then Exit Sub
  GLock = True
  Design.StatusBar1.Panels(3) = Oper
  Call GetCursorPos(Mpos)
  Mpos.Y = Mpos.Y + 1 '>>> não parece estar na posição correcta sem o +1
  Select Case Oper
   Case 0 '-------------------------------------espera de acção
     Call ClientToScreen(Plan.hWnd, Wpos)
     X = Mpos.X - Wpos.X
     Y = Mpos.Y - Wpos.Y
     If Desenho.DrawingIN(X, Y) Then
       If (X <> oldX) Or (Y <> oldY) Then
         RulersH.SetCursor X
         RulersV.SetCursor Y
         oldX = X: oldY = Y
         TxtX = "X:" & GetCurPos(X, Desenho.TamanhoX, Desenho.MeioX, Desenho.CentroX, Desenho.Zoom)
         TxtY = "Y:" & GetCurPos(Y, Desenho.TamanhoY, Desenho.MeioY, Desenho.CentroY, Desenho.Zoom)
       End If
      Else
       If oldX <> -10000 Then
         RulersH.HideCursor
         RulersV.HideCursor
         oldX = -10000: oldY = -10000
       End If
     End If
   Case 2 '--------------------------------------- desenhar parede
     Call ClientToScreen(Plan.hWnd, Wpos)
     X = Mpos.X - Wpos.X
     Y = Mpos.Y - Wpos.Y
     If Desenho.DrawingIN(X, Y) = False Then Desenho.SetDrawingIN X, Y
     If (X <> oldX) Or (Y <> oldY) Then
       GridX = GetCurPos(X, Desenho.TamanhoX, Desenho.MeioX, Desenho.CentroX, Desenho.Zoom)
       GridY = GetCurPos(Y, Desenho.TamanhoY, Desenho.MeioY, Desenho.CentroY, Desenho.Zoom)
       Call ToGrid(GridX, GridY)
       RulersH.SetCursor Desenho.MapToScreen(GridX, 1)
       RulersV.SetCursor Desenho.MapToScreen(GridY, 2)
       Paredes.Item(SelectP).X2 = GridX
       Paredes.Item(SelectP).Y2 = GridY
       DrawDesignPlan
       TxtX = Paredes.Item(SelectP).GetTextX
       TxtY = Paredes.Item(SelectP).GetTextY
       oldX = X: oldY = Y
     End If
   Case 5 '--------------------------------------- mexer Janela
     Call ClientToScreen(Plan.hWnd, Wpos)
     X = Mpos.X - Wpos.X
     Y = Mpos.Y - Wpos.Y
     If Desenho.DrawingIN(X, Y) Then
       GridX = Janelas.Item(SelectJ).Position
       Janelas.Item(SelectJ).ChangePos Desenho, X, Y
       If GridX <> Janelas.Item(SelectJ).Position Then DrawDesignPlan
     End If
  End Select
  GLock = False
End Sub

Private Sub Tools_Resize()
  If GLock = False Then
   GLock = True
   Zoom.top = Design.Tools.ScaleHeight - 130
   ScrMapa.top = Tools.ScaleHeight - ScrMapa.Height - 1
   DrawTool.Height = Zoom.top - 3
   GLock = False
  End If
End Sub

Private Sub Zoom_Change()
  GLock = True
  If Ready Then Plan.SetFocus
  SetObjectPos
  GLock = False
End Sub

