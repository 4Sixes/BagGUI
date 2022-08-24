Attribute VB_Name = "Transparent"
Private Declare Function FindWindow _
    Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong _
    Lib "user32" Alias "GetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong _
    Lib "user32" Alias "SetWindowLongA" _
    (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes _
    Lib "user32" _
    (ByVal hWnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public hWnd As Long

Sub MakeTransparent(frm As Object, TransparentValue As Integer)

Dim bytOpacity As Byte

'Control the opacity setting.
bytOpacity = TransparentValue

hWnd = FindWindow("ThunderDFrame", frm.Caption)
Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
Call SetLayeredWindowAttributes(hWnd, 0, bytOpacity, LWA_ALPHA)

End Sub
