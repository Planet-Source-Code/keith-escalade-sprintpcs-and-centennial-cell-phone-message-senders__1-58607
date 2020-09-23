VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl Command 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   ClientHeight    =   795
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1590
   DefaultCancel   =   -1  'True
   FillStyle       =   0  'Solid
   ScaleHeight     =   53
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   106
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   360
   End
   Begin PicClip.PictureClip pc 
      Left            =   120
      Top             =   360
      _ExtentX        =   2381
      _ExtentY        =   556
      _Version        =   393216
      Cols            =   5
      Picture         =   "Button.ctx":0000
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Command"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   885
   End
End
Attribute VB_Name = "Command"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Winux Graphic User Interface for Windows based systems
'Copyright (C) 2002-2003 Winux Team
'This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version.
'This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
'You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'A copy of this licence is available in root\system directory.
'http://www.winux.free.fr or tex_winux@hotmail.com for more details.

'Created by Teh Ming Han (teh_minghan@hotmail.com)

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINT_API) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINT_API) As Long

Public Enum State_b
    Normal_ = 0
    Default_ = 1
End Enum

Dim m_State As State_b
Dim m_Font As Font

Const m_Def_State = State_b.Normal_

Private Type POINT_API
    X As Long
    Y As Long
End Type

Dim s As Integer
Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseOut()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Private Sub lbl_Change()
    UserControl_Resize
End Sub

Private Sub lbl_Click()
    UserControl_Click
End Sub

Private Sub lbl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Timer1_Timer()
    Dim pnt As POINT_API
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt

    If pnt.X < UserControl.ScaleLeft Or _
       pnt.Y < UserControl.ScaleTop Or _
       pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
       pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
       
        Timer1.Enabled = False
        RaiseEvent MouseOut
        statevalue_pic
    End If
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    statevalue_pic
End Sub

Private Sub UserControl_InitProperties()
    state_value = m_Def_State
    Enabled = True
    Caption = Ambient.DisplayName
    Set Font = UserControl.Ambient.Font
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
    make_xpbutton 1
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Timer1.Enabled = True
    If X >= 0 And Y >= 0 And _
       X <= UserControl.ScaleWidth And Y <= UserControl.ScaleHeight Then
        RaiseEvent MouseMove(Button, Shift, X, Y)
        If Button = vbLeftButton Then
            make_xpbutton 1
        Else: make_xpbutton 3
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    statevalue_pic
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    state_value = PropBag.ReadProperty("State", m_Def_State)
    Enabled = PropBag.ReadProperty("Enabled", True)
    Caption = PropBag.ReadProperty("Caption", Ambient.DisplayName)
    Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    statevalue_pic
    If Enabled = True Then lbl.ForeColor = vbBlack Else lbl.ForeColor = RGB(161, 161, 146)
End Property

Private Sub UserControl_Resize()
    statevalue_pic
    lbl.Top = (UserControl.ScaleHeight - lbl.Height) / 2
    lbl.Left = (UserControl.ScaleWidth - lbl.Width) / 2
End Sub

Private Sub UserControl_Show()
    statevalue_pic
End Sub

Private Sub UserControl_Terminate()
    statevalue_pic
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("State", m_State, m_Def_State)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Caption", lbl.Caption, Ambient.DisplayName)
    Call PropBag.WriteProperty("Font", m_Font, UserControl.Ambient.Font)
End Sub

Public Property Get State() As State_b
Attribute State.VB_Description = "Returns/sets the state of the command button when mouse_out."
Attribute State.VB_ProcData.VB_Invoke_Property = ";Misc"
    State = m_State
End Property

Public Property Let State(ByVal vNewValue As State_b)
    m_State = vNewValue
    PropertyChanged "State"
    statevalue_pic
End Property

Private Sub statevalue_pic()
    If State = Default_ Then
        s = 4
    ElseIf State = Normal_ Then
        s = 0
    End If
    
    If UserControl.Enabled = True Then
        make_xpbutton s
    Else: make_xpbutton 2
    End If
End Sub

Private Sub make_xpbutton(z As Integer)
    UserControl.ScaleMode = 3 'Draw in pixels
    Dim brx, bry, bw, bh As Integer
    'Short cuts
    brx = UserControl.ScaleWidth - 3 'right x
    bry = UserControl.ScaleHeight - 3 'right y
    bw = UserControl.ScaleWidth - 6 'border width - corners width
    bh = UserControl.ScaleHeight - 6 'border height - corners height
    'Draws button
    'Goes clockwise first for corners(first four)
    'followed by borders(next four) and center(last step).
    UserControl.PaintPicture pc.GraphicCell(z), 0, 0, 3, 3, 0, 0, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, 0, 3, 3, 15, 0, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, bry, 3, 3, 15, 18, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), 0, bry, 3, 3, 0, 18, 3, 3
    UserControl.PaintPicture pc.GraphicCell(z), 3, 0, bw, 3, 3, 0, 12, 3
    UserControl.PaintPicture pc.GraphicCell(z), brx, 3, 3, bh, 15, 3, 3, 15
    UserControl.PaintPicture pc.GraphicCell(z), 0, 3, 3, bh, 0, 3, 3, 15
    UserControl.PaintPicture pc.GraphicCell(z), 3, bry, bw, 3, 3, 18, 12, 3
    UserControl.PaintPicture pc.GraphicCell(z), 3, 3, bw, bh, 3, 3, 12, 15

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Caption.VB_UserMemId = -518
    Caption = lbl.Caption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    lbl.Caption() = vNewCaption
    PropertyChanged "Caption"
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vNewFont As Font)
    Set m_Font = vNewFont
    Set UserControl.Font = vNewFont
    Set lbl.Font = m_Font
    Call UserControl_Resize
    PropertyChanged "Font"
End Property
