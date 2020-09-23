VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Menu"
   ClientHeight    =   165
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   1320
   LinkTopic       =   "Form2"
   ScaleHeight     =   165
   ScaleWidth      =   1320
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuShow 
         Caption         =   "Show"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
      Begin VB.Menu dash1 
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
Private Sub mnuExit_Click()
Unload Form1
End
End Sub
Private Sub mnuShow_Click()
Form1.Show
End Sub
