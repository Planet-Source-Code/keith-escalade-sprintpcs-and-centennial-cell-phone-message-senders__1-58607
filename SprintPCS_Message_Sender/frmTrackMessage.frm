VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Track a Message"
   ClientHeight    =   5415
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
   Icon            =   "frmTrackMessage.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5415
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dock"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox Text3 
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
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5040
      Width           =   5055
   End
   Begin SprintPCSMessageSender.PCS PCS1 
      Left            =   2400
      Top             =   600
      _ExtentX        =   1349
      _ExtentY        =   1005
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
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2175
   End
   Begin SprintPCSMessageSender.Command Command1 
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      Caption         =   "Track"
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   15
      Top             =   4200
      Width           =   3495
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
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
      TabIndex        =   14
      Top             =   4200
      Width           =   1455
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   240
      Y1              =   4800
      Y2              =   5040
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   360
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Message Sent"
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
      TabIndex        =   13
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   12
      Top             =   3720
      Width           =   3495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3240
      Width           =   3495
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   1680
      TabIndex        =   10
      Top             =   2760
      Width           =   3495
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
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
      TabIndex        =   9
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Received:"
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
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Time Sent:"
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
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   5040
      Y1              =   2040
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      X1              =   4920
      X2              =   840
      Y1              =   2040
      Y2              =   2040
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
      TabIndex        =   6
      Top             =   1920
      Width           =   735
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
      TabIndex        =   2
      Top             =   240
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   0
      X2              =   5280
      Y1              =   720
      Y2              =   720
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
      TabIndex        =   3
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tracking Number"
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
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   600
      X2              =   360
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      X1              =   360
      X2              =   240
      Y1              =   1200
      Y2              =   1440
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   0
      Picture         =   "frmTrackMessage.frx":0CCE
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FormDrag(WhatForm As Form)
ReleaseCapture
Call SendMessage(WhatForm.hwnd, &HA1, 2, 0&)
End Sub
Private Sub Check1_Click()
Select Case Check1.Value
Case Is = 0
Case Is = 1
If Form1.Visible = True Then Me.Top = Form1.Top: Me.Left = Form1.Left + Form1.Width
End Select
End Sub
Private Sub Command1_Click()
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label13.Caption = ""
Text2.Text = "Tracking Message..."
Text3.Text = ""
PCS1.TrackMessage Text1
End Sub
Private Sub Form_Load()
Shape1.Height = Me.Height
Shape1.Width = Me.Width
End Sub
Private Sub Label1_Click()
Me.Hide
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
If Check1.Value = 1 Then If Form1.Visible = True Then Form1.Top = Me.Top: Form1.Left = Me.Left - Form1.Width
End Sub
Private Sub PCS1_CouldntTrackMessage()
Text2.Text = "Error: Unable to track message."
End Sub
Private Sub PCS1_GotTrackingStatus(TheStatus As String, SentTime As String, ReceiveTime As String, FromSource As String, ToDestination As String, MessageSent As String)
Text2.Text = TheStatus
Label9.Caption = SentTime
Label10.Caption = ReceiveTime
Label11.Caption = FromSource
Label13.Caption = ToDestination
Text3.Text = MessageSent
End Sub
Private Sub PCS1_SockError(Description As String)
Text2.Text = "Error: Unable to track message."
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
