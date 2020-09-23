VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TJanela 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   2985
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNo 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   825
   End
   Begin VB.CommandButton CmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   330
      Left            =   90
      TabIndex        =   4
      Top             =   900
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aspect"
      Height          =   1500
      Left            =   1035
      TabIndex        =   0
      Top             =   90
      Width           =   1860
      Begin VB.TextBox JLeft 
         Height          =   330
         Left            =   810
         TabIndex        =   6
         Top             =   810
         Width           =   690
      End
      Begin VB.TextBox JWidth 
         Height          =   330
         Left            =   810
         TabIndex        =   3
         Top             =   270
         Width           =   690
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   1485
         TabIndex        =   2
         Top             =   270
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Position from left"
         Height          =   390
         Left            =   135
         TabIndex        =   7
         Top             =   765
         Width           =   660
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Width"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Image Image1 
      Height          =   690
      Left            =   135
      Picture         =   "TJanela.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   690
   End
End
Attribute VB_Name = "TJanela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BkWidth As Long
Public BkLeft As Long

Private Sub CmdNO_Click()
  JWidth = BkWidth
  JLeft = BkLeft
  TJanela.Hide
End Sub

Private Sub cmdOK_Click()
  TJanela.Hide
End Sub
