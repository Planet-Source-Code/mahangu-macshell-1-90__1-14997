VERSION 5.00
Begin VB.Form frmControl 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "MacSHell Control Panel"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picmain 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Label lblcap 
         BackStyle       =   0  'Transparent
         Caption         =   "MacSHELL Control Panel"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
      Begin VB.Image imglogo 
         Height          =   480
         Left            =   -120
         Picture         =   "frmControl.frx":0000
         ToolTipText     =   "Close this window!"
         Top             =   -120
         Width           =   480
      End
   End
   Begin VB.Label lblSysProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "System Properties"
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image imgSystemProperties 
      Height          =   480
      Left            =   1800
      Picture         =   "frmControl.frx":08CA
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label lblNetworkProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Network Properties"
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image imgNetworkProperties 
      Height          =   480
      Left            =   480
      Picture         =   "frmControl.frx":0D0C
      Top             =   1680
      Width           =   480
   End
   Begin VB.Label lblKeyboardProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyboard Properties"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgKeyboardProperties 
      Height          =   480
      Left            =   4320
      Picture         =   "frmControl.frx":114E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblDisplayProperties 
      BackStyle       =   0  'Transparent
      Caption         =   "Display Properties"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imgDisplayProperties 
      Height          =   480
      Left            =   3120
      Picture         =   "frmControl.frx":1590
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add New Hardware"
      Height          =   735
      Left            =   1680
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgAddHardware 
      Height          =   480
      Left            =   1920
      Picture         =   "frmControl.frx":19D2
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblAddRemove 
      BackStyle       =   0  'Transparent
      Caption         =   "Add / Remove - Programs"
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Image imgAddRemove 
      Height          =   480
      Left            =   480
      Picture         =   "frmControl.frx":1E14
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
picmain.Width = ScaleWidth
End Sub

Private Sub Form_LostFocus()
Unload Me

End Sub

Private Sub Form_Resize()
picmain.Width = ScaleWidth
End Sub

Private Sub imgAddHardware_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)

End Sub

Private Sub imgAddRemove_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)

End Sub

Private Sub imgDisplayProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)

End Sub

Private Sub imgKeyboardProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)

End Sub

Private Sub imglogo_Click()
Unload Me

End Sub

Private Sub imgNetworkProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)

End Sub

Private Sub imgSystemProperties_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)

End Sub

Private Sub lblcap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub picmain_Click()
FormMove Me
End Sub

Private Sub picmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub
