Attribute VB_Name = "SrtartUp"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'NEED THIS TO STAY ON TOP
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Sub Main()

frmSplash.Show
frmSplash.Refresh
Sleep 3000
Load Form1
Form1.Show
Unload frmSplash

End Sub

