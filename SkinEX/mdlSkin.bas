Attribute VB_Name = "mdlSkin"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_SYSCOMMAND = &H112
Public Const WM_MOVE = &HF012
Public Sub FormDrag(Form As Form)
Call ReleaseCapture
Call SendMessage(Form.hwnd, WM_SYSCOMMAND, WM_MOVE, 0)
End Sub
