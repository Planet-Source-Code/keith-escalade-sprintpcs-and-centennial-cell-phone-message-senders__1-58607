VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Quick Send"
   ClientHeight    =   1935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form6"
   ScaleHeight     =   1935
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   480
   End
   Begin SprintPCSMessageSender.PCS PCS1 
      Left            =   4560
      Top             =   840
      _ExtentX        =   1349
      _ExtentY        =   1005
   End
   Begin SprintPCSMessageSender.HzxYProgressBar PB1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      Max             =   3
      BarColorSet     =   2
      Bar_Pic         =   "frmQuickSend.frx":0000
      BorderColor     =   0
      BackColor       =   16777215
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
      TabIndex        =   2
      Top             =   360
      Width           =   5055
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
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   4800
      X2              =   840
      Y1              =   120
      Y2              =   120
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
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "frmQuickSend.frx":0142
      Top             =   1200
      Width           =   3960
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PNumber As String
Dim Message As String
Dim Callback As String
Private Sub Form_Load()
On Error Resume Next
If Command$ = "" Then Form5.Show: Unload Me: Exit Sub
Shape1.Height = Me.Height
Shape1.Width = Me.Width
FrmStayOnTop Me
PNumber = Split(Command$, "&&")(0)
Message = Split(Command$, "&&")(1)
Callback = Split(Command$, "&&")(2)
If IsNumeric(PNumber) = False Then Me.Hide: MsgBox "Invalid phone number", vbExclamation: End
If Callback <> "" Then If IsNumeric(Callback) = False Then Me.Hide: MsgBox "Invalid callback number", vbExclamation: End
If Len(Message) = 0 Then Me.Hide: MsgBox "Message must have a body", vbCritical: End
If Len(PNumber) <> 10 Then Me.Hide: MsgBox "PCS Number must be 10 characters long", vbCritical: End
If Len(Callback) <> 0 Then If Len(Callback) <> 10 Then Me.Hide: MsgBox "Callback number must be 10 characters long", vbCritical: End
PCS1.SendMessage PNumber, Message, Callback
PB1.Value = 1
Label1.Caption = "Connecting..."
End Sub
Private Sub PCS1_CouldntSendMessage()
PB1.Value = 0
Me.Hide
MsgBox "Couldn't send message to the PCS phone number", vbExclamation: End
End Sub
Private Sub PCS1_MessageSent(TrackingNumber As String)
PB1.Value = 3
Label1.Caption = "Message sent"
On Error Resume Next
Open App.Path & "\TrackingNumbers.log" For Append As #1
Print #1, TrackingNumber & vbCrLf & "     To: " & PNumber & vbCrLf & "     Message: " & Message & vbCrLf & "     Callback Number: " & Callback
Close #1
DoEvents
Timer1.Enabled = True
End Sub
Private Sub PCS1_ReceivingData()
Label1.Caption = "Receiving data..."
End Sub
Private Sub PCS1_SockError(Description As String)
PB1.Value = 0
Me.Hide
MsgBox "Error sending message:" & vbCrLf & Description, vbExclamation: End
End Sub
Private Sub PCS1_SockConnect()
PB1.Value = 2
Label1.Caption = "Connected... Sending data"
End Sub
Private Sub Timer1_Timer()
End
End Sub
