VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "MacShell 1.0.0"
   ClientHeight    =   5445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8415
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5445
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraCommand 
      BackColor       =   &H00FF8080&
      Caption         =   "Shell Execute Command"
      Height          =   855
      Left            =   3960
      TabIndex        =   12
      Top             =   960
      Width           =   2895
      Begin VB.TextBox txtCommand 
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.PictureBox picmain 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   0
      Width           =   1215
      Begin VB.Label lblmymac 
         BackStyle       =   0  'Transparent
         Caption         =   "My Mac"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblrun 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Run"
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblPowerOff 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Power Off"
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblfind 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Find"
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Line ln1 
         X1              =   0
         X2              =   1200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblclose 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Exit Shell"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "Main"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         ToolTipText     =   "Use the Main menu to get to the most common functions of MacSHELL."
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Label lblDeskContro 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Control Panel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   15
      Top             =   1080
      Width           =   855
   End
   Begin VB.Image imgControl 
      Height          =   480
      Left            =   1320
      Picture         =   "frmMain.frx":08CA
      Top             =   600
      Width           =   480
   End
   Begin VB.Label lblScrLock 
      BackStyle       =   0  'Transparent
      Caption         =   "Sreen Lock"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Width           =   855
   End
   Begin VB.Image imgScreenLock 
      Height          =   480
      Left            =   240
      Picture         =   "frmMain.frx":0D0C
      Top             =   4560
      Width           =   480
   End
   Begin VB.Label lbldeskmed 
      BackStyle       =   0  'Transparent
      Caption         =   "Media Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Width           =   615
   End
   Begin VB.Image imgdeskmed 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":114E
      ToolTipText     =   "Media Player - Listen to your favourite audio files."
      Top             =   3360
      Width           =   480
   End
   Begin VB.Label lbldeskfind 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   375
   End
   Begin VB.Image imgfind 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":1A18
      ToolTipText     =   "Find - Find what you wan't, fast and easily."
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label lbldeskrun 
      BackStyle       =   0  'Transparent
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   375
   End
   Begin VB.Image imgrun 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":22E2
      ToolTipText     =   "Run - Select a file or appilcation and run it."
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lbldeskmymac 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "My Mac"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Image imglogo 
      Height          =   480
      Left            =   -120
      Picture         =   "frmMain.frx":2BAC
      ToolTipText     =   "Click here to exit MacSHELL!"
      Top             =   -120
      Width           =   480
   End
   Begin VB.Image imgmymac 
      Height          =   480
      Left            =   360
      Picture         =   "frmMain.frx":3476
      ToolTipText     =   "My Mac - Explore your computer!"
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lbltop 
      BackColor       =   &H00FF8080&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Dir1_Change()
File1.Path = Dir1.Path

End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive


End Sub

Private Sub File1_Click()
File1.FileName
End Sub

Private Sub Command1_Click()
dlgcol.ShowColor
Call SetCol

End Sub

Private Sub Form_Click()
Call CloseMenus



End Sub


Private Sub imgControl_Click()
Call CloseMenus
frmControl.Show
End Sub

Private Sub imgdeskmed_Click()
frmMed.Show

End Sub

Private Sub imgfind_Click()
Call ShowFindDialog

End Sub

Private Sub imglogo_Click()
Call CloseMenus


EndApp



End Sub

Private Sub imgsd_Click()
End

End Sub



Private Sub imgmymac_Click()
Call CloseMenus
frmBrowse.Show


End Sub

Private Sub imgrun_Click()
Call ShowRunDialog(Me, "MacSHELL", _
        "Select the file you want to open.")


End Sub

Private Sub imgScreenLock_Click()
frmScreenLock.Show

End Sub

Private Sub lblclose_Click()
Call CloseMenus
Call EndApp



End Sub

Private Sub lblfind_Click()
Call ShowFindDialog
Call CloseMenus

End Sub

Private Sub lblmain_Click()
While picmain.Height <> "1445"
picmain.Height = picmain.Height + 1
Wend

End Sub

Private Sub lblmymac_Click()
Call CloseMenus
frmBrowse.Show

End Sub

Private Sub lblPowerOff_Click()
Call ShutDown
End Sub

Private Sub lblrun_Click()
Call ShowRunDialog(Me, "MacSHELL", _
        "Select the file you want to open.")

End Sub

Private Sub lbltop_Click()
Call CloseMenus
frmAbout.Show
End Sub

Private Sub lbltop_DblClick()
frmAbout.Show

End Sub

Private Sub picmain_Click()
While picmain.Height <> "1445"
picmain.Height = picmain.Height + 1
Wend

End Sub


Private Sub picmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
While picmain.Height <> "1445"
picmain.Height = picmain.Height + 1
Wend

End Sub

Private Sub txtCommand_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next
If KeyCode = 13 Then

    Shell txtCommand.Text, vbNormalFocus
         
    
End If

If KeyCode = 13 And txtCommand.Text = "exit" Then

    End
    
End If
End Sub
