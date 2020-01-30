VERSION 5.00
Begin VB.Form Doublons 
   Caption         =   "Doublons"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdannuler 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox listdoublons 
      Height          =   1035
      ItemData        =   "doublons.frx":0000
      Left            =   240
      List            =   "doublons.frx":0002
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Image apercu 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label lbldoublons 
      Caption         =   "Plusieurs occurences du même nom, sélectionner la bonne personne dans la liste :"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "Doublons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
Dim req As String
Dim rst As New ADODB.Recordset

req = "SELECT nom_perso,prenom_perso,nom_image WHERE DOUBLER.num_perso = PERSO.num_perso AND DOUBLER.num_image = IMAGE.num_image AND nom_image = " & "'" & apercu.Tag & "'"
'Seiyuu.cbperso.Text = rst!nom_perso & rst!prenom_perso
'Seiyuu.imgperso.Picture = apercu.Picture
Doublons.Hide
Seiyuu.SetFocus

'Seiyuu.cbperso.List (Seiyuu.cbperso.ListIndex)

'Seiyuu.listperso.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

cmdannuler_Click

End Sub

Private Sub listdoublons_Click()

chemin = "C:\WINNT\Profiles\Administrateur\Bureau\Projet\"
Select Case listdoublons.ListIndex
    Case 0
        apercu.Picture = LoadPicture(chemin & "ranmafille.bmp")
        apercu.Tag = "ranmafille.bmp"
    Case 1
        apercu.Picture = LoadPicture(chemin & "ranmagarçon.bmp")
        apercu.Tag = "ranmagarçon.bmp"
End Select
apercu.ToolTipText = "Visuel de " & Seiyuu.cbperso.Text

End Sub

Private Sub cmdannuler_Click()

Doublons.Hide

End Sub

Private Sub listdoublons_DblClick()

cmdok_Click

End Sub
