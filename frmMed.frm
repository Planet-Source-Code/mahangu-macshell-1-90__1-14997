VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMed 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "MacSHELL Media Player"
   ClientHeight    =   1500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgopen 
      Left            =   2160
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "MacSHELL Media Player"
      Filter          =   "Mp3 Files | *.mp3"
      InitDir         =   "C:\"
   End
   Begin VB.PictureBox picCap 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Label lblstate 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading"
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblcap 
         BackStyle       =   0  'Transparent
         Caption         =   "MacSHELL Media Player --->"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   0
         Width           =   2055
      End
      Begin VB.Image imglogo 
         Height          =   480
         Left            =   -120
         Picture         =   "frmMed.frx":0000
         Top             =   -120
         Width           =   480
      End
   End
   Begin VB.Image imgopen 
      Height          =   480
      Left            =   120
      Picture         =   "frmMed.frx":08CA
      Top             =   240
      Width           =   480
   End
   Begin MediaPlayerCtl.MediaPlayer medmain 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "frmMed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_LostFocus()
Unload Me

End Sub

Private Sub imglogo_Click()
Unload Me

End Sub

Private Sub imgopen_Click()
dlgopen.ShowOpen
medmain.FileName = dlgopen.FileName


End Sub

Private Sub lblcap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub lblstate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub

Private Sub medmain_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
If medmain.PlayState = mpPlaying Then lblstate.Caption = "Playing"
If medmain.PlayState = mpStopped Then lblstate.Caption = "Stopped"
If medmain.PlayState = mpWaiting Then lblstate.Caption = "Waiting"
If medmain.PlayState = mpClosed Then lblstate.Caption = "Closed"
If medmain.PlayState = mpPaused Then lblstate.Caption = "Paused"
If medmain.PlayState = mpScanForward Then lblstate.Caption = "Scanning Forward"
If medmain.PlayState = mpScanReverse Then lblstate.Caption = "Scanning Reverse"






End Sub

Private Sub picCap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormMove Me
End Sub
