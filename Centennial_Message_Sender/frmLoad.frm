VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Centennial Message Sender"
   ClientHeight    =   1050
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoad.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   1050
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   240
      Top             =   240
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
      Caption         =   "Message Sender"
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
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Picture         =   "frmLoad.frx":0CCE
      Top             =   0
      Width           =   3840
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Shape1.Width = Me.Width
Shape1.Height = Me.Height
FrmStayOnTop Me
End Sub
Private Sub Timer1_Timer()
Me.Hide
Form1.Show
Timer1.Enabled = False
End Sub

