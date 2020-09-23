VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ForeColor       =   &H00FFC0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox piccap 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Image imglogo 
         Height          =   480
         Left            =   -120
         Picture         =   "frmAbout.frx":0000
         Top             =   -120
         Width           =   480
      End
      Begin VB.Label lblCap 
         BackStyle       =   0  'Transparent
         Caption         =   "About MacSHELL"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Label lblcontact 
      BackStyle       =   0  'Transparent
      Caption         =   "To get the FREE source code for MacSHELL visit www.planet-source-code.com/vb"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label lblrel 
      BackStyle       =   0  'Transparent
      Caption         =   "Release Date : 2001/1/3"
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblver 
      BackStyle       =   0  'Transparent
      Caption         =   "Version : 1.9.0 (Version 2 Beta 1)"
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lbldesc 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":08CA
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin VB.Image imgAbout 
      Height          =   480
      Left            =   120
      Picture         =   "frmAbout.frx":09A5
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imglogo_Click()
Unload Me

End Sub

Private Sub lblcap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me

End Sub

Private Sub picCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub
