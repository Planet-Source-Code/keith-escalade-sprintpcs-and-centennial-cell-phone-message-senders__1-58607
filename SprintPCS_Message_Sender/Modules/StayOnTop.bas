Attribute VB_Name = "StayOnTop"
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Sub FrmStayOnTop(WhatForm As Form)
Call SetWindowPos(WhatForm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
Public Sub FrmDontStayOnTop(WhatForm As Form)
Call SetWindowPos(WhatForm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
End Sub
