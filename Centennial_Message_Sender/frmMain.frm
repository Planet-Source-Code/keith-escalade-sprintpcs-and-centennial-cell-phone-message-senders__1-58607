VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Centennial Message Sender"
   ClientHeight    =   6255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin CentennialMessageSender.Centennial Centennial1 
      Left            =   4320
      Top             =   960
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin CentennialMessageSender.HzxYProgressBar PB1 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5760
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      Max             =   3
      BarColorSet     =   2
      Bar_Pic         =   "frmMain.frx":0CCE
   End
   Begin CentennialMessageSender.Command Command1 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "Send"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   1365
      Left            =   120
      MaxLength       =   110
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   2280
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
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
      Height          =   285
      Left            =   120
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Garren Fitzenreiter 2005"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   5160
      Width           =   4095
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   240
      Y1              =   4320
      Y2              =   4560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   360
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "10-Digit Callback Number (Not Required)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   4200
      Width           =   3975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "110 Characters remaining"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3720
      Width           =   4575
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   240
      Y1              =   2040
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   360
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Message To Send"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "10-Digit Centennial Phone Number"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   240
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   5280
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5295
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Picture         =   "frmMain.frx":0E10
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public DontShow As Boolean
Dim HoldNumber As String
Dim HoldMessage As String
Dim HoldCallback As String
Private Sub Centennial1_CouldntSendMessage()
PB1.Value = 0
Status.Caption = "Couldn't send message to the Centennial phone number"
End Sub
Private Sub Centennial1_MessageSent()
PB1.Value = 3
Status.Caption = "Message sent"
On Error Resume Next
Open App.Path & "\SendLog.log" For Append As #1
Print #1, Now & vbCrLf & "     To: " & HoldNumber & vbCrLf & "     Message: " & HoldMessage & vbCrLf & "     Callback Number: " & HoldCallback
Close #1
End Sub
Private Sub Centennial1_ReceivingData()
Status.Caption = "Receiving data..."
End Sub
Private Sub Centennial1_SockConnect()
PB1.Value = 2
Status.Caption = "Connected... Sending data"
End Sub
Private Sub Centennial1_SockError(Description As String)
PB1.Value = 0
Status.Caption = "Error"
End Sub
Private Sub Command1_Click()
If Len(Text2.Text) = 0 Then Status.Caption = "Message must have a body": Exit Sub
If Len(Text1.Text) <> 10 Then Status.Caption = "Centennial Number must be 10 characters long": Exit Sub
If Len(Text3.Text) <> 0 Then If Len(Text3.Text) <> 10 Then Status.Caption = "Callback number must be 10 characters long": Exit Sub
Centennial1.SendMessage Text1, Text2, Text3
HoldNumber = Text1
HoldMessage = Text2
HoldCallback = Text3
Status.Caption = "Connecting..."
PB1.Value = 1
End Sub
Private Sub Form_Load()
Shape1.Height = Me.Height
Shape1.Width = Me.Width
End Sub
Private Sub Label1_Click()
End
End Sub
Private Sub Label2_Click()
If DontShow = True Then Me.Hide: Exit Sub
If DontShow = False Then Form3.Show: Me.Hide: Exit Sub
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub
Private Sub FormDrag(WhatForm As Form)
ReleaseCapture
Call SendMessage(WhatForm.hWnd, &HA1, 2, 0&)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
Case 7755: 'Right click icon event
PopupMenu Form2.mnuMenu
Case 7725: 'Double click icon event
Me.Show
End Select
End Sub
Private Sub Form_Initialize()
sIcon.cbSize = Len(sIcon)
sIcon.hWnd = Me.hWnd
sIcon.uId = vbNull
sIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
sIcon.uCallBackMessage = WM_MOUSEMOVE
sIcon.hIcon = Me.Icon
sIcon.szTip = "Centennial Message Sender" & vbNullChar
Call Shell_NotifyIcon(NIM_ADD, sIcon)
Call Shell_NotifyIcon(NIM_MODIFY, sIcon)
End Sub
Private Sub Text1_GotFocus()
Text1.BackColor = vbWhite
Text1.ForeColor = vbBlack
End Sub
Private Sub Text1_LostFocus()
Text1.BackColor = &H404040
Text1.ForeColor = vbWhite
End Sub
Private Sub Text2_GotFocus()
Text2.BackColor = vbWhite
Text2.ForeColor = vbBlack
End Sub
Private Sub Text2_LostFocus()
Text2.BackColor = &H404040
Text2.ForeColor = vbWhite
End Sub
Private Sub Text3_GotFocus()
Text3.BackColor = vbWhite
Text3.ForeColor = vbBlack
End Sub
Private Sub Text3_LostFocus()
Text3.BackColor = &H404040
Text3.ForeColor = vbWhite
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
Private Sub Text2_Change()
Label6.Caption = 110 - Len(Text2) & " Characters remaining"
End Sub
Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then Exit Sub
If IsNumeric(Chr(KeyAscii)) = False Then KeyAscii = 0
End Sub
