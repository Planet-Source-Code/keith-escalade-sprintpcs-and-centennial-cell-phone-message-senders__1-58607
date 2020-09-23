Attribute VB_Name = "SysTray"
Public Type NOTIFYICONDATA
cbSize As Long
hwnd As Long
uId As Long
uFlags As Long
uCallBackMessage As Long
hIcon As Long
szTip As String * 64
End Type
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, ptry As NOTIFYICONDATA) As Boolean
Public sIcon As NOTIFYICONDATA


'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Select Case X
'Case 7755: 'Right click icon event
'
'Case 7725: 'Double click icon event
'
'End Select
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Sub Form_Initialize()
'sIcon.cbSize = Len(sIcon)
'sIcon.hwnd = Me.hwnd
'sIcon.uId = vbNull
'sIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'sIcon.uCallBackMessage = WM_MOUSEMOVE
'sIcon.hIcon = Me.Icon
'sIcon.szTip = "keith_escalade" & vbNullChar
'Call Shell_NotifyIcon(NIM_ADD, sIcon)
'Call Shell_NotifyIcon(NIM_MODIFY, sIcon)
'End Sub
