VERSION 5.00
Begin VB.Form Splash 
   BackColor       =   &H00413E01&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H00413E01&
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   7020
      Begin VB.PictureBox picLogo 
         AutoSize        =   -1  'True
         Height          =   2010
         Left            =   120
         Picture         =   "Splash.frx":0000
         ScaleHeight     =   1950
         ScaleWidth      =   6690
         TabIndex        =   1
         Top             =   240
         Width           =   6750
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00413E01&
         Caption         =   "v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   240
         TabIndex        =   4
         Tag             =   "1039"
         Top             =   2280
         Width           =   510
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         BackColor       =   &H00413E01&
         Caption         =   "Windows 32Bits"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   2400
         TabIndex        =   3
         Tag             =   "1040"
         Top             =   2280
         Width           =   2205
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00413E01&
         Caption         =   "Manuel Augusto Nogueira dos Santos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Tag             =   "1043"
         Top             =   2640
         Width           =   6735
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    LoadResStrings Me
    
End Sub

