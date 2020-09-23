Attribute VB_Name = "FormDrag"
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
'Private Sub FormDrag(WhatForm As Form)
'Rem Makes a form draggable. Ex: control_mousedown()
'Rem                                              FormDrag me
'ReleaseCapture
'Call SendMessage(WhatForm.hwnd, &HA1, 2, 0&)
'End Sub

