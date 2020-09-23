VERSION 5.00
Begin VB.Form Pattern 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   1680
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "Pattern.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   109
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   112
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   11
      Left            =   1260
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   13
      Top             =   855
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   10
      Left            =   1260
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   12
      Top             =   450
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   9
      Left            =   1260
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   11
      Top             =   45
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   8
      Left            =   855
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   10
      Top             =   855
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   7
      Left            =   855
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   9
      Top             =   450
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   6
      Left            =   855
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   8
      Top             =   45
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   5
      Left            =   450
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   7
      Top             =   855
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   4
      Left            =   450
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   6
      Top             =   450
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   3
      Left            =   450
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   5
      Top             =   45
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   2
      Left            =   45
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   4
      Top             =   855
      Width           =   360
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   1
      Left            =   45
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   3
      Top             =   450
      Width           =   360
   End
   Begin VB.CommandButton CmdNO 
      Caption         =   "Cancel"
      Height          =   330
      Left            =   900
      TabIndex        =   2
      Top             =   1260
      Width           =   735
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "Ok"
      Height          =   330
      Left            =   45
      TabIndex        =   1
      Top             =   1260
      Width           =   735
   End
   Begin VB.PictureBox Desenhos 
      AutoRedraw      =   -1  'True
      Height          =   360
      Index           =   0
      Left            =   45
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   20
      TabIndex        =   0
      Top             =   45
      Width           =   360
   End
End
Attribute VB_Name = "Pattern"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdNO_Click()
  Pattern.Hide
End Sub

Private Sub cmdOk_Click()
  Pattern.Hide
End Sub
Private Sub DrawDesenhos()
  Dim i As Integer
  
  For i = 0 To 11
    If i = DesP Then
      Pattern.Desenhos(DesP).Appearance = 0
      Pattern.Desenhos(DesP).BackColor = RGB(0, 50, 50)
      Pattern.Desenhos(i).FillColor = RGB(255, 255, 100)
    Else
      Pattern.Desenhos(i).Appearance = 1
      Pattern.Desenhos(i).BackColor = RGB(50, 50, 0)
      Pattern.Desenhos(i).FillColor = RGB(10, 155, 150)
    End If
    Pattern.Desenhos(i).FillStyle = Dbox(i)
    Pattern.Desenhos(i).Line (0, 0)-Step(24, 24), , B
  Next i

End Sub

Private Sub Desenhos_Click(Index As Integer)
  DesP = Index
  DrawDesenhos
End Sub

Private Sub Form_Activate()
  DrawDesenhos
End Sub
