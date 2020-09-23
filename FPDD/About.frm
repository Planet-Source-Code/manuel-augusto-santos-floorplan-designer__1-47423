VERSION 5.00
Begin VB.Form About 
   BackColor       =   &H0030280F&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Project1"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   ClipControls    =   0   'False
   FillColor       =   &H0032290F&
   FontTransparent =   0   'False
   ForeColor       =   &H0032290F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1052"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5160
      TabIndex        =   0
      Tag             =   "1054"
      Top             =   2280
      Width           =   1467
   End
   Begin VB.Image Image1 
      Height          =   1950
      Left            =   0
      Picture         =   "About.frx":0000
      Top             =   120
      Width           =   6690
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808000&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6720
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    LoadResStrings Me
    
End Sub


Private Sub cmdOK_Click()
        Unload Me
End Sub

