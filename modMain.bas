Attribute VB_Name = "modMain"
Sub CloseMenus()
While frmMain.picmain.Height <> "255"
frmMain.picmain.Height = frmMain.picmain.Height - 1
Wend
End Sub

Sub EndApp()
Unload frmMain
Unload frmBrowse
End
End Sub


