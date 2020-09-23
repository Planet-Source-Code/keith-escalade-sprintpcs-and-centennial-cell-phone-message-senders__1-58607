VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "System Tray"
   ClientHeight    =   1575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3930
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   1575
   ScaleWidth      =   3930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Don't tell me again"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin SprintPCSMessageSender.Command Command1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      caption         =   "OK"
      font            =   "frmSysTray.frx":0CCE
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   255
      Left            =   0
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The program has been minimized to the system tray in the bottom right corner of your screen."
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
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Left            =   960
      Picture         =   "frmSysTray.frx":0CF6
      Top             =   600
      Width           =   1860
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Check1.Value = 1 Then Form1.DontShow = True: Me.Hide
If Check1.Value = 0 Then Form1.DontShow = False: Me.Hide
End Sub
Private Sub Form_Load()
FrmStayOnTop Me
Shape1.Height = Me.Height
Shape1.Width = Me.Width
End Sub
