VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmBrowse 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.DriveListBox drv 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "res://D:\WINNT\System32\shdoclc.dll/offcancl.htm#http:///"
   End
   Begin VB.PictureBox picmain 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      Begin VB.Label lblcap 
         BackStyle       =   0  'Transparent
         Caption         =   "My Mac "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   0
         Width           =   2535
      End
      Begin VB.Image imglogo 
         Height          =   480
         Left            =   -120
         Picture         =   "frmBrowse.frx":0000
         ToolTipText     =   "Close this window!"
         Top             =   -120
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmBrowse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub drv_Change()
web.Navigate drv.Drive

End Sub

Private Sub Form_Load()
web.Navigate "C:\"
picmain.Width = frmBrowse.Width


End Sub

Private Sub Form_Resize()
web.Height = ScaleHeight
web.Width = ScaleWidth
picmain.Width = frmBrowse.Width


End Sub

Private Sub imglogo_Click()
Unload frmBrowse

End Sub
