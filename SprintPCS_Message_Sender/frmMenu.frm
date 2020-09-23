VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Menu Form"
   ClientHeight    =   45
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   810
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   45
   ScaleWidth      =   810
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Collapse Menu"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label6_Click()
Form1.Hide
End Sub
Private Sub mnuExit_Click()
Unload Form1
End
End Sub
Private Sub mnuShow_Click()
Form1.Show
End Sub
