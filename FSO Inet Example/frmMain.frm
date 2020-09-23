VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Your Program"
   ClientHeight    =   5415
   ClientLeft      =   2790
   ClientTop       =   2295
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   6585
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuRegister 
         Caption         =   "Register"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuRegister_Click()
frmRegister.Show
End Sub
