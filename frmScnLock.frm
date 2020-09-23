VERSION 5.00
Begin VB.Form frmScreenLock 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "MacSHELL - ScreenLock"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Unlock"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmScreenLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogin_Click()
If txtUser.Text = UN And txtPass.Text = Pass Then
    Unload Me
End If

End Sub

Private Sub Form_Load()
Dim UN
Dim Pass

UN = InputBox("Please enter the Username you want to use.", "Screenlock Setup - Username")
Pass = InputBox("Please enter the Password you want to use.", "Screenlock Setup - Password")

DisableCAD True


End Sub

Private Sub Form_Terminate()
DisableCAD = True

End Sub
