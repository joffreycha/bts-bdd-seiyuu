VERSION 5.00
Begin VB.Form Visuel 
   Caption         =   "Aperçu"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4515
   LinkTopic       =   "Form2"
   ScaleHeight     =   4170
   ScaleWidth      =   4515
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Visuel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name)
Visuel.Hide

End Sub
