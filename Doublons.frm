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
   Begin VB.ListBox listnumdoub 
      Height          =   645
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox listimgdoub 
      Height          =   450
      Left            =   3720
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdannuler 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin VB.ListBox listdoublons 
      Height          =   1035
      ItemData        =   "Doublons.frx":0000
      Left            =   240
      List            =   "Doublons.frx":0002
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Image dblapercu 
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label lbldoublons 
      BackStyle       =   0  'Transparent
      Caption         =   "Plusieurs occurences du même nom, sélectionnez la bonne personne dans la liste :"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   3735
   End
End
Attribute VB_Name = "Doublons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force la déclaration des variables

Dim rst As New ADODB.Recordset
Private Sub cmdok_Click()

rst.Open "SELECT nom_seiyuu,prenom_seiyuu,nom_images FROM SEIYUU,DOUBLER,IMAGES WHERE DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND num_perso = '" & listnumdoub.List(listnumdoub.ListIndex) & "' AND nom_images = '" & dblapercu.Tag & "'", Seiyuu.cn, adOpenDynamic, adLockPessimistic

'nom du doubleur correspondant au personnage
Seiyuu.cbseiyuu.Text = rst!nom_seiyuu & " " & rst!prenom_seiyuu

rst.Close
Unload Doublons
Seiyuu.SetFocus

End Sub

Private Sub Form_Load()

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name) 'changement du thème

End Sub

Private Sub listdoublons_Click()
'charge une image en fonction de l'élément sélectionné

listnumdoub.ListIndex = listdoublons.ListIndex
dblapercu.Picture = LoadPicture()
On Error Resume Next
dblapercu.Picture = LoadPicture(Command & "Persos\" & listimgdoub.List(listdoublons.ListIndex))
If dblapercu.Picture = 0 Then
    dblapercu.Picture = LoadPicture(cheminperso & "Persos\" & "non_disponible.jpg")
End If
On Error GoTo 0
dblapercu.Tag = listimgdoub.List(listdoublons.ListIndex)
dblapercu.ToolTipText = "Aperçu de " & Seiyuu.cbperso.Text

End Sub

Private Sub cmdannuler_Click()

Unload Doublons 'décharge la form

End Sub

Private Sub listdoublons_DblClick()

cmdok_Click

End Sub
