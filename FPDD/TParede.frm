VERSION 5.00
Begin VB.Form TParede 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   195
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   168
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdNo 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1350
      TabIndex        =   5
      Top             =   2475
      Width           =   1050
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   2475
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Caption         =   "Type"
      Height          =   1140
      Left            =   90
      TabIndex        =   3
      Top             =   1260
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aspect"
      Height          =   1140
      Left            =   855
      TabIndex        =   0
      Top             =   45
      Width           =   1545
      Begin VB.ComboBox ComboWidth 
         Height          =   315
         ItemData        =   "TParede.frx":0000
         Left            =   855
         List            =   "TParede.frx":0013
         TabIndex        =   2
         Text            =   "20"
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Lespessura 
         Alignment       =   1  'Right Justify
         Caption         =   "Width"
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   315
         Width           =   735
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "TParede.frx":002B
      Top             =   315
      Width           =   480
   End
End
Attribute VB_Name = "TParede"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BkWidth As Long

Private Sub CmdNO_Click()
  ComboWidth.Text = BkWidth
  TParede.Hide
End Sub

Private Sub cmdOK_Click()
  TParede.Hide
End Sub

