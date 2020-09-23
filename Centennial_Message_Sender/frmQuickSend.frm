VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Quick Send"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   LinkTopic       =   "Form4"
   ScaleHeight     =   1935
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CentennialMessageSender.Centennial Centennial1 
      Left            =   2400
      Top             =   720
      _ExtentX        =   979
      _ExtentY        =   979
   End
   Begin CentennialMessageSender.HzxYProgressBar PB1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      Max             =   3
      BarColorSet     =   2
      Bar_Pic         =   "frmQuickSend.frx":0000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3960
      Top             =   1320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   4920
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   840
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
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
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Picture         =   "frmQuickSend.frx":0142
      Top             =   1080
      Width           =   3840
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PNumber As String
Dim Message As String
Dim CallBack As String
Private Sub Form_Load()
On Error Resume Next
If Command$ = "" Then Form5.Show: Unload Me: Exit Sub
Shape1.Height = Me.Height
Shape1.Width = Me.Width
FrmStayOnTop Me
PNumber = Split(Command$, "&&")(0)
Message = Split(Command$, "&&")(1)
CallBack = Split(Command$, "&&")(2)
If IsNumeric(PNumber) = False Then Me.Hide: MsgBox "Invalid phone number", vbExclamation: End
If CallBack <> "" Then If IsNumeric(CallBack) = False Then Me.Hide: MsgBox "Invalid callback number", vbExclamation: End
If Len(Message) = 0 Then Me.Hide: MsgBox "Message must have a body", vbCritical: End
If Len(PNumber) <> 10 Then Me.Hide: MsgBox "PCS Number must be 10 characters long", vbCritical: End
If Len(CallBack) <> 0 Then If Len(CallBack) <> 10 Then Me.Hide: MsgBox "Callback number must be 10 characters long", vbCritical: End
Centennial1.SendMessage PNumber, Message, CallBack
PB1.Value = 1
Label1.Caption = "Connecting..."
End Sub
Private Sub Centennial1_CouldntSendMessage()
PB1.Value = 0
Me.Hide
MsgBox "Couldn't send message to the Centennial phone number", vbExclamation: End
End Sub
Private Sub Centennial1_MessageSent()
PB1.Value = 3
Label1.Caption = "Message sent"
On Error Resume Next
Open App.Path & "\SendLog.log" For Append As #1
Print #1, Now & vbCrLf & "     To: " & PNumber & vbCrLf & "     Message: " & Message & vbCrLf & "     Callback Number: " & CallBack
Close #1
DoEvents
Timer1.Enabled = True
End Sub
Private Sub Centennial1_ReceivingData()
Label1.Caption = "Receiving data..."
End Sub
Private Sub Centennial1_SockError(Description As String)
PB1.Value = 0
Me.Hide
MsgBox "Error sending message:" & vbCrLf & Description, vbExclamation: End
End Sub
Private Sub Centennial1_SockConnect()
PB1.Value = 2
Label1.Caption = "Connected... Sending data"
End Sub
Private Sub Timer1_Timer()
End
End Sub

