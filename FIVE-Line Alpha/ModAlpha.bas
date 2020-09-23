Attribute VB_Name = "ModAlpha"
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function Layer Lib "user32" Alias "SetLayeredWindowAttributes" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Function SetLayered(hWnd As Long, Volume As Long)
SetWindowLong hWnd, (-20), &H80000: Layer hWnd, 0, CByte(Volume), &H2   'Set Alpha
End Function
