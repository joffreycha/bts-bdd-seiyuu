VERSION 5.00
Begin VB.Form Seiyuu 
   BackColor       =   &H80000004&
   Caption         =   "Doubleurs de japanimation"
   ClientHeight    =   9525
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10980
   ForeColor       =   &H80000008&
   Icon            =   "Seiyuu.frx":0000
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   9525
   ScaleWidth      =   10980
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optgenerique 
      Caption         =   "Gé&nérique de fin"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   18
      ToolTipText     =   "Choix du générique"
      Top             =   600
      Width           =   1815
   End
   Begin VB.OptionButton optgenerique 
      Caption         =   "G&énérique de début"
      Height          =   255
      Index           =   0
      Left            =   5760
      TabIndex        =   17
      ToolTipText     =   "Choix du générique"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CheckBox chkmedia 
      Caption         =   "Afficher le &lecteur"
      Height          =   255
      Left            =   360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      ToolTipText     =   "Affiche les commandes du lecteur"
      Top             =   8880
      Width           =   1575
   End
   Begin VB.PictureBox picreel 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   4680
      ScaleHeight     =   2055
      ScaleWidth      =   2055
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdportrait 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Por&trait"
      Height          =   255
      Index           =   1
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Voir un portrait du personnage"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.ListBox listperso 
      Height          =   1620
      ItemData        =   "Seiyuu.frx":030A
      Left            =   720
      List            =   "Seiyuu.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   6720
      Width           =   2895
   End
   Begin VB.ComboBox cbperso 
      Height          =   315
      ItemData        =   "Seiyuu.frx":030E
      Left            =   720
      List            =   "Seiyuu.frx":0310
      Sorted          =   -1  'True
      TabIndex        =   9
      ToolTipText     =   "Recherche par personnage"
      Top             =   5880
      Width           =   2895
   End
   Begin VB.CommandButton cmdportrait 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Po&rtrait"
      Height          =   255
      Index           =   0
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Voir un portrait du doubleur"
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox listseiyuu 
      Height          =   1620
      ItemData        =   "Seiyuu.frx":0312
      Left            =   720
      List            =   "Seiyuu.frx":0314
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   3120
      Width           =   2895
   End
   Begin VB.ComboBox cbseiyuu 
      Height          =   315
      ItemData        =   "Seiyuu.frx":0316
      Left            =   720
      List            =   "Seiyuu.frx":0318
      Sorted          =   -1  'True
      TabIndex        =   5
      ToolTipText     =   "Recherche par doubleur"
      Top             =   2280
      Width           =   2895
   End
   Begin VB.ComboBox cbtitre 
      Height          =   315
      ItemData        =   "Seiyuu.frx":031A
      Left            =   2520
      List            =   "Seiyuu.frx":0324
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.ComboBox cbtheme 
      Height          =   315
      ItemData        =   "Seiyuu.frx":0342
      Left            =   13320
      List            =   "Seiyuu.frx":0352
      Style           =   2  'Dropdown List
      TabIndex        =   16
      ToolTipText     =   "Changement du thème de l'application"
      Top             =   240
      Width           =   1695
   End
   Begin VB.ComboBox cbanime 
      Height          =   315
      ItemData        =   "Seiyuu.frx":0382
      Left            =   2520
      List            =   "Seiyuu.frx":0384
      Sorted          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Recherche par série"
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblpersoseiyuu 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Personna&ges incarnés par le doubleur :"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Label lblseiyuuanime 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Dou&bleurs de la série :"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Image imgapercu 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   1
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label lblperso 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&Personnages"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   5400
      Width           =   975
   End
   Begin VB.Shape shpperso 
      Height          =   3015
      Left            =   360
      Top             =   5640
      Width           =   7095
   End
   Begin VB.Label lblseiyuu 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&Doubleurs"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Image imgapercu 
      BorderStyle     =   1  'Fixed Single
      Height          =   2055
      Index           =   0
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Shape shpseiyuu 
      Height          =   3015
      Left            =   360
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label lbltitre 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "Sélection p&ar :"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lbltheme 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "T&hème de bureau :"
      Height          =   255
      Left            =   11880
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblanime 
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "&Série :"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   495
   End
   Begin VB.Menu Mfichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Mouvrir 
         Caption         =   "&Ouvrir"
      End
      Begin VB.Menu Mfermer 
         Caption         =   "&Fermer"
      End
      Begin VB.Menu tiret 
         Caption         =   "-"
      End
      Begin VB.Menu Mquitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu Moptions 
      Caption         =   "&Options"
      Begin VB.Menu Majouter 
         Caption         =   "&Ajouter"
      End
      Begin VB.Menu Msupprimer 
         Caption         =   "&Supprimer"
      End
      Begin VB.Menu Mmodifier 
         Caption         =   "&Modifier"
      End
      Begin VB.Menu Mrechercher 
         Caption         =   "&Rechercher"
      End
   End
End
Attribute VB_Name = "Seiyuu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Version 1.7 version EPTI
Option Explicit 'force la déclaration des variables
Option Compare Text 'ne prend pas en compte la casse

Public cn As New ADODB.Connection 'connexion ADODB
Public modif As Integer '0 -> création
                        '1 -> modification d'un personnage
                        '2 -> modification d'un seiyuu
                        '3 -> modification d'un anime
Public idperso As Integer 'num_perso
Public idseiyuu As Integer 'num_seiyuu
Public idanime As Integer 'num_anime
Public idimages As Integer 'num_images
Dim rst As New ADODB.Recordset 'recordset
Dim doubl As Integer 'test de doublon
Dim focusseiyuu As String
Dim focusperso As String
Dim titre As String
Dim recherche As Integer '0 -> anime ; 1 -> seiyuu ; 2 -> personnage

Public Function recup_titre(anime As String) As String
'récupère le titre du générique de fin ou de début de l'anime entré en paramètre

Dim rst2 As New ADODB.Recordset 'nouveau jeu d'enregistrement

rst2.Open "SELECT generique_deb,generique_fin FROM ANIME WHERE (titre_original= '" & anime & "'Or titre_version_fr = '" & anime & "')", cn, adOpenForwardOnly, adLockPessimistic

If Seiyuu.optgenerique(0).Value = True Then
    If rst2!generique_deb <> "" Then recup_titre = rst2!generique_deb
Else
    If rst2!generique_fin <> "" Then recup_titre = rst2!generique_fin
End If

End Function

Private Sub cbanime_GotFocus()

If cn.State = 1 And cbanime.Text <> "" Then
    Mmodifier.Enabled = True
    Msupprimer.Enabled = True
    modif = 3
End If

End Sub

Private Sub cbanime_LostFocus()

Msupprimer.Enabled = False
Mmodifier.Enabled = False

End Sub

Private Sub cbtheme_Click()
'change le thème

Call choix_theme(cbtheme.ListIndex, Me.Name)
Call choix_theme(cbtheme.ListIndex, "Lecteur")
Call choix_theme(cbtheme.ListIndex, "Visuel")

End Sub

Private Sub cbtitre_Click()
'recherche par titre original ou version française

Dim titreanime As String 'titre de l'anime

If cn.State = 1 Then
    titreanime = cbanime.Text
    cbanime.Clear
    
    'remplissage de cbanime
    If cbtitre.ListIndex = 0 Then 'titre original
        rst.Open "SELECT titre_original FROM ANIME", cn, adOpenForwardOnly, adLockReadOnly
        While Not rst.EOF
            cbanime.AddItem rst!titre_original
            rst.MoveNext
        Wend
    ElseIf cbtitre.ListIndex = 1 Then 'titre version française
        rst.Open "SELECT titre_version_fr FROM ANIME", cn, adOpenForwardOnly, adLockReadOnly
        While Not rst.EOF
            cbanime.AddItem rst!titre_version_fr
            rst.MoveNext
        Wend
    End If
    rst.Close
    If titreanime <> "" Then listperso_Click
End If

End Sub

Private Sub chkmedia_Click()
'affiche ou masque le lecteur

If chkmedia.Value = 0 Then
    Lecteur.Hide
Else
    Lecteur.Show
End If

End Sub

Private Sub imgapercu_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'affiche l'image sélectionnée dans sa taille d'origine

If Index = 0 Then 'photo du seiyuu
    picreel.Picture = imgapercu(0)
    picreel.Top = 2280
    picreel.Visible = True
    imgapercu(Index).Visible = False
Else 'image du personnage
    picreel.Picture = imgapercu(1)
    picreel.Top = 5880
    picreel.Visible = True
    imgapercu(Index).Visible = False
End If

End Sub

Private Sub imgapercu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'réaffiche l'image sélectionnée dans l'ImageBox

picreel.Picture = LoadPicture()
picreel.Visible = False
imgapercu(Index).Visible = True

End Sub

Private Sub listperso_GotFocus()

Mmodifier.Enabled = True
Msupprimer.Enabled = True
modif = 1

End Sub

Private Sub listperso_LostFocus()

Mmodifier.Enabled = False
Msupprimer.Enabled = False

End Sub

Private Sub listseiyuu_DblClick()

cbseiyuu_Click

End Sub

Private Sub listseiyuu_GotFocus()

Mmodifier.Enabled = True
Msupprimer.Enabled = True
modif = 2

End Sub

Private Sub listseiyuu_LostFocus()

Mmodifier.Enabled = False
Msupprimer.Enabled = False

End Sub

Private Sub Majouter_Click()
'lance la form ajouter

modif = 0
Load Ajouter
Ajouter.Show vbModal

End Sub

Public Sub Mfermer_Click()
'initialise les contrôles

'ListBox
listseiyuu.Visible = False
listperso.Visible = False
listseiyuu.Clear
listperso.Clear

'ComboBox
cbperso.Text = ""
cbseiyuu.Text = ""
cbanime.Text = ""
cbanime.Clear
cbseiyuu.Clear
cbperso.Clear

'CommandButton
cmdportrait(0).Enabled = False
cmdportrait(1).Enabled = False

'Menu
Mouvrir.Enabled = True
Mfermer.Enabled = False
Majouter.Enabled = False
Msupprimer.Enabled = False
Mmodifier.Enabled = False
Mrechercher.Enabled = False

'Label
lblseiyuuanime.Visible = False
lblpersoseiyuu.Visible = False

'ImageBox
imgapercu(0).Picture = LoadPicture()
imgapercu(1).Picture = LoadPicture()
imgapercu(0).ToolTipText = ""
imgapercu(1).ToolTipText = ""

Lecteur.WMP.URL = ""

Unload Visuel

recherche = 0

'fermeture de la connexion
If cn.State = 1 Then
    cn.Close
End If

End Sub

Private Sub Mmodifier_Click()
'lance la feuille "modifier"

Load Ajouter
Ajouter.Show vbModal

End Sub

Public Sub Mouvrir_Click()
'ouvre la connexion à la BD et remplit les ComboBox

cn.Open "EPTI", "ig05", "ig05" 'ouverture de la connexion

'remplissage de cbanime
If cbtitre.ListIndex = 0 Then
    rst.Open "SELECT titre_original FROM ANIME", cn, adOpenForwardOnly, adLockReadOnly
    While Not rst.EOF
        cbanime.AddItem rst!titre_original
        rst.MoveNext
    Wend
ElseIf cbtitre.ListIndex = 1 Then
    rst.Open "SELECT titre_version_fr FROM ANIME", cn, adOpenForwardOnly, adLockReadOnly
    While Not rst.EOF
        cbanime.AddItem rst!titre_version_fr
        rst.MoveNext
    Wend
End If
rst.Close

'remplissage de cbseiyuu
rst.Open "SELECT nom_seiyuu,prenom_seiyuu FROM SEIYUU", cn, adOpenForwardOnly, adLockReadOnly
While Not rst.EOF
    If Not IsNull(rst!nom_seiyuu) Then
        If Not IsNull(rst!prenom_seiyuu) Then
            cbseiyuu.AddItem rst!nom_seiyuu & " " & rst!prenom_seiyuu
        Else
            cbseiyuu.AddItem rst!nom_seiyuu
        End If
    ElseIf IsNull(rst!nom_seiyuu) And Not IsNull(rst!prenom_seiyuu) Then
            cbseiyuu.AddItem rst!prenom_seiyuu
    End If
    rst.MoveNext
Wend
rst.Close

'remplissage de cbperso
rst.Open "SELECT nom_perso,prenom_perso FROM PERSO", cn, adOpenForwardOnly, adLockReadOnly
While Not rst.EOF
    If Not IsNull(rst!nom_perso) Then
        If Not IsNull(rst!prenom_perso) Then
            cbperso.AddItem rst!nom_perso & " " & rst!prenom_perso
        Else
            cbperso.AddItem rst!nom_perso
        End If
    ElseIf IsNull(rst!nom_perso) And Not IsNull(rst!prenom_perso) Then
            cbperso.AddItem rst!prenom_perso
    End If
    rst.MoveNext
Wend
rst.Close

Mouvrir.Enabled = False
Mfermer.Enabled = True
Majouter.Enabled = True
Mrechercher.Enabled = True

End Sub

Private Sub Mquitter_Click()

'fermeture de la connexion
If cn.State = 1 Then
    cn.Close
End If

'déchargement des feuilles
Unload Doublons
Unload Lecteur
Unload Visuel

End 'fermeture du programme

End Sub

Private Sub Mrechercher_Click()
'lance la form rechercher

Rechercher.Show vbModal 'empêche l'utilisateur d'agir sur une autre fenêtre

End Sub

Private Sub Form_Unload(Cancel As Integer)

Mquitter_Click

End Sub

Private Sub listperso_DblClick()

cbperso_Click

End Sub

Public Sub cbperso_Click()

cbperso_KeyPress (13) 'pression sur la touche "Entrée"

End Sub

Private Sub cbperso_KeyPress(KeyAscii As Integer)
'recherche par personnage

Dim i As Integer 'indice
Dim req As String 'requête SQL
Dim req2 As String 'requête SQL
Dim rst2 As New ADODB.Recordset 'recordset
Dim tablo() As String 'tableau contenant le nom et le prenom

If KeyAscii = 13 Then 'pression sur la touche "Entrée"
    For i = 0 To cbperso.ListCount - 1
        If cbperso.Text = "" Then
            cbperso.ListIndex = -1
            Exit Sub
        ElseIf cbperso.List(i) Like cbperso.Text Then
            tablo = Split(cbperso.List(i), " ")
            
            If UBound(tablo) = 0 Then 'seulement un nom ou un prénom
                req = "SELECT PERSO.num_perso,titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_images FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND (nom_perso = " & "'" & tablo(0) & "' OR prenom_perso = '" & tablo(0) & "')"
            Else 'nom et prénom
                req = "SELECT PERSO.num_perso,titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_images FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = '" & tablo(1) & "'"
            End If
            
            rst.Open req, cn, adOpenDynamic, adLockPessimistic
            
            On Error GoTo absent
            
            'titre de l'anime correspondant au personnage sélectionné
            If cbtitre.ListIndex = 0 Then
                cbanime.Text = rst!titre_original
            ElseIf cbtitre.ListIndex = 1 Then
               cbanime.Text = rst!titre_version_fr
            End If
            
            'éléments sur lequel il faut mettre le focus
            focusperso = cbperso.Text
            focusseiyuu = rst!nom_seiyuu & " " & rst!prenom_seiyuu

            'remplissage de la liste des doublons
            doubl = 0
            If UBound(tablo) = 0 Then
                While Not rst.EOF
                    If cbperso.List(i) = rst!nom_perso Then
                        doubl = doubl + 1
                        Doublons.listdoublons.AddItem rst!nom_perso
                        Doublons.listimgdoub.AddItem rst!nom_images
                        Doublons.listnumdoub.AddItem rst!num_perso
                    ElseIf cbperso.List(i) = rst!prenom_perso Then
                        Doublons.listdoublons.AddItem rst!prenom_perso
                        Doublons.listimgdoub.AddItem rst!nom_images
                        Doublons.listnumdoub.AddItem rst!num_perso
                    End If
                    rst.MoveNext
                Wend
            Else
                While Not rst.EOF
                    If cbperso.List(i) = rst!nom_perso & " " & rst!prenom_perso Then
                        doubl = doubl + 1
                        Doublons.listdoublons.AddItem rst!nom_perso & " " & rst!prenom_perso
                        Doublons.listimgdoub.AddItem rst!nom_images
                        Doublons.listnumdoub.AddItem rst!num_perso
                    End If
                    rst.MoveNext
                Wend
            End If
            rst.Close
            
            recherche = 2 'recherche par personnage
            
            'gestion des doublons de personnages
            If doubl > 1 Then
                Doublons.listdoublons.ListIndex = 0
                Load Doublons
                Doublons.Show vbModal 'empêche l'utilisateur d'agir sur une autre fenêtre
                cbseiyuu_Click
            Else
                Unload Doublons
                cbanime_Click
            End If
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucun personnage ne correspond au nom saisi.", vbOKOnly + vbExclamation, "Erreur saisie personnage"
    If cbperso.Text <> "" And cn.State = 0 Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Connexion fermée"
    Exit Sub
    
absent:
    MsgBox "Il n'y a pas de série ou de doubleur correspondant à ce personnage !", vbCritical + vbOKOnly, "Erreur personnage"
    rst.Close
End If

End Sub

Public Sub cbseiyuu_Click()

cbseiyuu_KeyPress (13) 'pression sur la touche "Entrée"

End Sub

Private Sub cbanime_Click()

cbanime_KeyPress (13) 'pression sur la touche "Entrée"

End Sub

Private Sub cbanime_KeyPress(KeyAscii As Integer)
'recherche par anime

Dim req As String 'requête SQL
Dim i As Integer 'indice
Dim j As Integer 'indice

If KeyAscii = 13 Then 'pression sur la touche "Entrée"
    
    'affichage des listes lors de la saisie d'un titre d'anime
    For i = 0 To cbanime.ListCount - 1
        If cbanime.Text = "" Then
            cbanime.ListIndex = -1
            Exit Sub
            
        ElseIf cbanime.List(i) Like cbanime.Text Then
            rst.Open "SELECT titre_original,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,generique_deb,generique_fin FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND (titre_original = '" & cbanime.Text & "' OR titre_version_fr = '" & cbanime.Text & "')", cn, adOpenDynamic, adLockPessimistic
            
            'remplissage de la liste des seiyuu
            listseiyuu.Clear
            listseiyuu.Visible = True
            lblseiyuuanime.Visible = True
            j = 0
            While Not rst.EOF
                If Not IsNull(rst!nom_seiyuu) Then
                    If Not IsNull(rst!prenom_seiyuu) Then
                        listseiyuu.AddItem rst!nom_seiyuu & " " & rst!prenom_seiyuu
                    Else
                        listseiyuu.AddItem rst!nom_seiyuu
                    End If
                ElseIf IsNull(rst!nom_seiyuu) And Not IsNull(rst!prenom_seiyuu) Then
                    listseiyuu.AddItem rst!prenom_seiyuu
                End If
                rst.MoveNext
                Call supp_doublons(j, listseiyuu) 'suppression des doublons
            Wend
            rst.Close
            
            titre = cbanime.Text 'titre de l'anime sélectionné
           
            'focus sur le bon élément de la liste
            Select Case recherche
                Case 0 'recherche par anime
                    If listseiyuu.ListCount <> 0 Then
                        listseiyuu.ListIndex = 0
                    Else
                        MsgBox "Aucun doubleur dans cette série !", vbOKOnly + vbInformation, "Doubleur"
                    End If
                Case 1 'recherche par seiyuu
                    For j = 0 To listseiyuu.ListCount - 1
                        If listseiyuu.List(j) = focusseiyuu Then
                            listseiyuu.ListIndex = j
                            Exit For
                        End If
                    Next j
                Case 2 'recherche par personnage
                    For j = 0 To listseiyuu.ListCount - 1
                        If listseiyuu.List(j) = focusseiyuu Then
                            listseiyuu.ListIndex = j
                            Exit For
                        End If
                    Next j
            End Select
                    
            recherche = 0
            
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucune série ne correspond au titre saisi.", vbOKOnly + vbExclamation, "Erreur saisie série"
    If cbanime.Text <> "" And cn.State = 0 Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Connexion fermée"
End If

End Sub

Public Sub choix_theme(numtheme As Integer, nomform As String)
'change le thème de la form entrée en paramètre

Select Case numtheme
    Case 0
        Call change_theme(nomform, &H4000&, &H8000&, &H80000004, "3dwarro.cur", "vert")
    Case 1
        Call change_theme(nomform, &H400000, &HFFC0C0, &H80000004, "harrow.cur", "bleu")
    Case 2
        Call change_theme(nomform, &H0&, &HC0&, &H80000004, "3dwarro.cur", "rouge")
    Case 3
        Call change_theme(nomform, &H80000004, &H80000004, &H0&, "3dwarro.cur", "noir")
End Select

End Sub

Private Sub cmdportrait_Click(Index As Integer)
'affiche le portrait d'un personnage
'Index = photo sélectionnée

If Index = 0 Then
    If imgapercu(0).Tag = Command & "Seiyuu\non_disponible.jpg" Then
        MsgBox "Aucune image disponible !", vbOKOnly + vbCritical, "Portrait"
    Else
        Visuel.Picture = imgapercu(0).Picture
        Visuel.Show
    End If
ElseIf Index = 1 Then
    If imgapercu(1).Tag <> "" Then
        Visuel.Picture = LoadPicture(Command & "Portraits\" & imgapercu(1).Tag)
        Visuel.Show
    Else
        MsgBox "Aucune image disponible !", vbOKOnly + vbCritical, "Portrait"
    End If
End If

End Sub

Private Sub Form_Load()

recherche = 0

cbtheme.ListIndex = 0 'initialisation du thème

cbtitre.ListIndex = 0
Majouter.Enabled = False
Msupprimer.Enabled = False
Mmodifier.Enabled = False
picreel.Visible = False
chkmedia.Value = 0 'masque le lecteur par défaut
optgenerique(0).Value = True 'générique de début par défaut
Mfermer_Click 'initialise les contrôles

End Sub

Private Sub listperso_Click()
'image et titre de l'anime par rapport au personnage sélectionné

Dim req As String 'requête SQL
Dim tablo() As String 'tableau contenant le nom et le prenom du personnage
Dim tablo2() As String 'tableau contenant le nom et le prenom du seiyuu
Dim j As Integer 'indice

tablo = Split(listperso.List(listperso.ListIndex), " ")
tablo2 = Split(listseiyuu.List(listseiyuu.ListIndex), " ")
imgapercu(1).Picture = LoadPicture()
imgapercu(1).Tag = ""

If UBound(tablo) = 0 Then 'seulement un nom ou un prénom
    req = "SELECT PERSO.num_perso,SEIYUU.num_seiyuu,ANIME.num_anime,IMAGES.num_images,titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_images,portrait FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = '" & tablo2(0) & "' AND prenom_seiyuu = '" & tablo2(1) & "' AND (nom_perso = " & "'" & tablo(0) & "' Or prenom_perso = " & "'" & tablo(0) & "')"
Else 'nom et prénom
    req = "SELECT PERSO.num_perso,SEIYUU.num_seiyuu,ANIME.num_anime,IMAGES.num_images,titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_images,portrait FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = '" & tablo2(0) & "' AND prenom_seiyuu = '" & tablo2(1) & "' AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = " & "'" & tablo(1) & "'"
End If

rst.Open req, cn, adOpenDynamic, adLockPessimistic

'titre correspondant au personnage sélectionné
If cbtitre.ListIndex = 0 Then
    cbanime.Text = rst!titre_original
ElseIf cbtitre.ListIndex = 1 Then
    cbanime.Text = rst!titre_version_fr
End If
    
'vérification de l'existence de l'image
On Error Resume Next
imgapercu(1).Picture = LoadPicture(Command & "Persos\" & rst!nom_images)
If imgapercu(1).Picture = 0 Then
    imgapercu(1).Picture = LoadPicture(Command & "Persos\" & "non_disponible.jpg")
End If
On Error GoTo 0

'sauvegarde du nom du portrait
If Not IsNull(rst!portrait) Then imgapercu(1).Tag = rst!portrait
    
'récupération des identifiants pour modification et suppression
idperso = rst!num_perso
idseiyuu = rst!num_seiyuu
idanime = rst!num_anime
idimages = rst!num_images

rst.Close

imgapercu(1).ToolTipText = "image de " & listperso.List(listperso.ListIndex)
cmdportrait(1).Enabled = True
cbperso.Text = listperso.List(listperso.ListIndex)

End Sub

Private Sub listseiyuu_Click()
'liste des personnages et photo par rapport au seiyuu sélectionné

Dim tablo() As String 'tableau contenant le nom et le prenom
Dim j As Integer 'indice
Dim i As Integer 'indice

tablo = Split(listseiyuu.List(listseiyuu.ListIndex), " ")
rst.Open "SELECT titre_original,nom_seiyuu,prenom_seiyuu,photo_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = " & "'" & tablo(0) & "' AND prenom_seiyuu = " & "'" & tablo(1) & "'", cn, adOpenDynamic, adLockPessimistic

imgapercu(0).Picture = LoadPicture()

'vérification de l'existence de la photo
On Error Resume Next
imgapercu(0).Picture = LoadPicture(Command & "Seiyuu\" & rst!photo_seiyuu)
If imgapercu(0).Picture = 0 Then
    imgapercu(0).Picture = LoadPicture(Command & "Seiyuu\" & "non_disponible.jpg")
End If
On Error GoTo 0

'remplissage de la liste des personnages
listperso.Clear
listperso.Visible = True
lblpersoseiyuu.Visible = True
j = 0
While Not rst.EOF
    If Not IsNull(rst!nom_perso) Then
        If Not IsNull(rst!prenom_perso) Then
            listperso.AddItem rst!nom_perso & " " & rst!prenom_perso
        Else
            listperso.AddItem rst!nom_perso
        End If
    ElseIf IsNull(rst!nom_perso) And Not IsNull(rst!prenom_perso) Then
        listperso.AddItem rst!prenom_perso
    End If
    rst.MoveNext
    Call supp_doublons(j, listperso) 'suppression des doublons
Wend
rst.Close

'focus sur un personnage correspondant à l'anime et au seiyuu sélectionnés
If recherche = 0 Or recherche = 1 Then
    rst.Open "SELECT titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = '" & tablo(0) & "' AND prenom_seiyuu = '" & tablo(1) & "' AND (titre_original = '" & titre & "' Or titre_version_fr = '" & titre & "')", cn, adOpenDynamic, adLockPessimistic
    For i = 0 To listperso.ListCount
        If (listperso.List(i) = rst!nom_perso & " " & rst!prenom_perso) Or (listperso.List(i) = rst!nom_perso) Or (listperso.List(i) = rst!prenom_perso) Then
            rst.Close
            Exit For
       End If
    Next i
End If

imgapercu(0).ToolTipText = "photo de " & listseiyuu.List(listseiyuu.ListIndex)
cmdportrait(0).Enabled = True
cbseiyuu.Text = listseiyuu.List(listseiyuu.ListIndex)
If recherche = 0 Or recherche = 1 Then
    listperso.ListIndex = i 'focus sur un personnage de l'anime doublé par le seiyuu
ElseIf recherche = 2 Then
    For i = 0 To listperso.ListCount - 1
        If listperso.List(i) = focusperso Then
            listperso.ListIndex = i
            Exit For
        End If
    Next i
End If

End Sub

Private Sub cbseiyuu_KeyPress(KeyAscii As Integer)
'recherche par seiyuu

Dim req As String 'requête SQL
Dim req2 As String 'requête SQL
Dim i As Integer 'indice
Dim j As Integer 'indice
Dim tablo() As String 'tableau contenant le nom et le prenom

If KeyAscii = 13 Then 'pression sur la touche "Entrée"

    'affichage des listes lors de la saisie du nom d'un seiyuu
    For i = 0 To cbseiyuu.ListCount - 1
        If cbseiyuu.Text = "" Then
            cbseiyuu.ListIndex = -1
            Exit Sub
        ElseIf cbseiyuu.List(i) Like cbseiyuu.Text Then
            tablo = Split(cbseiyuu.List(i), " ")
            
            rst.Open "SELECT titre_original,titre_version_fr,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_images,portrait FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = " & "'" & tablo(0) & "' AND prenom_seiyuu = " & "'" & tablo(1) & "' ORDER BY nom_perso,prenom_perso", cn, adOpenDynamic, adLockPessimistic
            
            On Error GoTo absent
            
            'titre correspondant au seiyuu sélectionné
            If cbtitre.ListIndex = 0 Then
                cbanime.Text = rst!titre_original
            ElseIf cbtitre.ListIndex = 1 Then
                cbanime.Text = rst!titre_version_fr
            End If
            rst.Close
            
            focusseiyuu = cbseiyuu.Text 'élément sur lequel il faut mettre le focus
            
            'si ce n'est pas une recherche pas personnage
            If recherche <> 2 Then recherche = 1 'recherche par seiyuu
            
            cbanime_Click
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucun doubleur ne correspond au nom saisi.", vbOKOnly + vbExclamation, "Erreur saisie doubleur"
    If cbseiyuu.Text <> "" And cn.State = 0 Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Connexion fermée"
    Exit Sub
    
absent:
    MsgBox "Il n'y a pas de série ou de personnage correspondant à ce doubleur !", vbCritical + vbOKOnly, "Erreur doubleur"
    rst.Close

End If

End Sub

Private Sub change_theme(nomform As String, coulfond As String, coulbouton As String, coulpolice As String, curseur As String, picbouton As String)
'changement du thème de l'application

Dim i As Integer 'indice

Select Case nomform
    Case "Ajouter"
        'Ajouter
        With Ajouter
            'Form
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
            
            'OptionButton
            For i = 0 To .optchoix.Count - 1
                .optchoix(i).BackColor = coulbouton
                .optchoix(i).ForeColor = coulpolice
            Next i
            
            'Label
            .lblanime.ForeColor = coulpolice
            .lblnomperso.ForeColor = coulpolice
            .lblnomseiyuu.ForeColor = coulpolice
            .lblprenomperso.ForeColor = coulpolice
            .lblprenomseiyuu.ForeColor = coulpolice
            
            'CommandButton
            .cmdannuler.BackColor = coulbouton
            .cmdok.BackColor = coulbouton
         End With
         
    Case "Doublons"
        'Doublons
        With Doublons
            'Form
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
    
            'Label
            .lbldoublons.ForeColor = coulpolice
    
            'CommandButton
            .cmdannuler.BackColor = coulbouton
            .cmdok.BackColor = coulbouton
        End With

    Case "Explorateur"
        'Explorateur
        With Explorateur
            'Form
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
    
            'CommandButton
            .cmdannuler.BackColor = coulbouton
            .cmdenregistrer.BackColor = coulbouton
            .cmdouvrir.BackColor = coulbouton
    
            'Label
            .lblenregistrer.ForeColor = coulpolice
            .lblnomfic.ForeColor = coulpolice
            .lbltype.ForeColor = coulpolice
        End With

    Case "Lecteur"
        'Lecteur
        With Lecteur
            'Form
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
    
            'CommandButton
            .cmdretour.Picture = LoadPicture(Command & "Boutons\" & picbouton & "_retour.jpg")
            If .WMP.playState = wmppsPlaying Then
                .cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & picbouton & "_pause.jpg")
            Else
                .cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & picbouton & "_play.jpg")
            End If
            .cmdstop.Picture = LoadPicture(Command & "Boutons\" & picbouton & "_stop.jpg")
            .cmdavance.Picture = LoadPicture(Command & "Boutons\" & picbouton & "_avance.jpg")
            .cmdplus.BackColor = coulbouton
            .cmdmoins.BackColor = coulbouton
            .cmdnouveau.BackColor = coulbouton
            .cmdouvrir.BackColor = coulbouton
            .cmdsauver.BackColor = coulbouton
            .cmdfermer.BackColor = coulbouton
            .cmdplaypause.Tag = picbouton
    
            'CheckBox
            .chkvolume.BackColor = coulbouton
    
            'Label
            .lblduree.ForeColor = coulpolice
            .lbltitre.ForeColor = coulpolice
        End With

    Case "Rechercher"
        'Rechercher
        With Rechercher
            'Form
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
    
            'OptionButton
            For i = 0 To .optprenom.Count - 1
                .optprenom(i).BackColor = coulbouton
                .optprenom(i).ForeColor = coulpolice
            Next i
       
            'CommandButton
            .cmdfermer.BackColor = coulbouton
            .cmdrechercher.BackColor = coulbouton
        End With

    Case "Seiyuu"
        'Seiyuu
        With Seiyuu
            'Form
            .MouseIcon = LoadPicture(Command & "Curseurs\" & curseur)
            .Picture = LoadPicture(Command & "Wallpapers\" & cbtheme.List(cbtheme.ListIndex) & ".jpg")
    
            'OptionButton
            For i = 0 To .optgenerique.Count - 1
                .optgenerique(i).BackColor = coulbouton
                .optgenerique(i).ForeColor = coulpolice
            Next i
    
            'Label
            .lblanime.ForeColor = coulpolice
            .lblseiyuu.ForeColor = coulpolice
            .lblperso.ForeColor = coulpolice
            .lbltitre.ForeColor = coulpolice
            .lbltheme.ForeColor = coulpolice
            .lblpersoseiyuu.ForeColor = coulpolice
            .lblseiyuuanime.ForeColor = coulpolice
    
            'Shape
            .shpseiyuu.BorderColor = coulbouton
            .shpperso.BorderColor = coulbouton
    
            'CommandButton
            .cmdportrait(0).BackColor = coulbouton
            .cmdportrait(1).BackColor = coulbouton
    
            'CheckBox
            .chkmedia.BackColor = coulbouton
            .chkmedia.ForeColor = coulpolice
        End With

    Case "Visuel"
        'Visuel
        Visuel.BackColor = coulfond

End Select

End Sub

Private Sub supp_doublons(indice As Integer, controle As Control)
'traitement des doublons

If indice <> 0 Then
    If controle.List(indice) = controle.List(indice - 1) Then
        Doublons.listdoublons.AddItem controle.List(indice)
        controle.RemoveItem (indice)
    Else
        indice = indice + 1
    End If
Else
    indice = indice + 1
End If

End Sub

Private Sub Msupprimer_Click()
'procédure supprimer

Select Case modif
    Case 1 'suppression d'un personnage
        If MsgBox("Etes-vous sûr de vouloir supprimer ce personnage ?", vbYesNo + vbQuestion, "Suppression personnage") = vbYes Then
            
            'execution de la procédure stockée
            'paramètres : num_perso,num_images
            rst.Open "supprimer_perso '" & idperso & "','" & idimages & "'", cn, adOpenKeyset, adLockPessimistic
        End If
        
    Case 2 'suppression d'un seiyuu
        If MsgBox("Etes-vous sûr de vouloir supprimer ce doubleur ?", vbYesNo + vbQuestion, "Suppression doubleur") = vbYes Then
        
            'execution de la procédure stockée
            'paramètres : num_seiyuu
            rst.Open "supprimer_seiyuu '" & idseiyuu & "'"
        End If
    Case 3 'suppression d'un anime
        If MsgBox("Etes-vous sûr de vouloir supprimer cette série ?", vbYesNo + vbQuestion, "Suppression série") = vbYes Then
        
            'execution de la procédure stockée
            'paramètres : num_anime
            rst.Open "supprimer_anime '" & idanime & "'"
        End If
End Select

'mise à jour des ComboBox
Mfermer_Click
Mouvrir_Click

End Sub

