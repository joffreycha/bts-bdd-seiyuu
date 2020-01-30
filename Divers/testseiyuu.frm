VERSION 5.00
Begin VB.Form Seiyuu 
   BackColor       =   &H80000004&
   Caption         =   "Seiyuu"
   ClientHeight    =   6150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   Icon            =   "testseiyuu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   9600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdouvrir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Ouvrir"
      Height          =   375
      Left            =   7800
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Ouvrir la base de données"
      Top             =   4680
      Width           =   1335
   End
   Begin VB.ComboBox cbtheme 
      Height          =   315
      ItemData        =   "testseiyuu.frx":030A
      Left            =   7800
      List            =   "testseiyuu.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.ComboBox cbtitre 
      Height          =   315
      ItemData        =   "testseiyuu.frx":0324
      Left            =   3360
      List            =   "testseiyuu.frx":0331
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.PictureBox photo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   4680
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Aperçu"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.CommandButton cmdfermer 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   7800
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Fermer la connexion"
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdquitter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   7800
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Quitter"
      Top             =   5640
      Width           =   1335
   End
   Begin VB.ComboBox cbanime 
      Height          =   315
      ItemData        =   "testseiyuu.frx":035D
      Left            =   2520
      List            =   "testseiyuu.frx":035F
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
   Begin VB.Frame frmperso 
      BackColor       =   &H80000004&
      Caption         =   "&Perso"
      Height          =   2415
      Left            =   480
      TabIndex        =   7
      Top             =   3600
      Width           =   6615
      Begin VB.CommandButton cmdagrandir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Agran&dir"
         Height          =   255
         Index           =   1
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Voir l'image dans sa taille d'origine"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox listperso 
         Height          =   1425
         ItemData        =   "testseiyuu.frx":0361
         Left            =   360
         List            =   "testseiyuu.frx":0363
         Sorted          =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cbperso 
         Height          =   315
         ItemData        =   "testseiyuu.frx":0365
         Left            =   360
         List            =   "testseiyuu.frx":0367
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.Image imgperso 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame frmseiyuu 
      BackColor       =   &H80000004&
      Caption         =   "&Seiyuu"
      Height          =   2415
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   6615
      Begin VB.CommandButton cmdagrandir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "A&grandir"
         Height          =   255
         Index           =   0
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Voir l'image dans sa taille d'origine"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ListBox listseiyuu 
         Height          =   1425
         ItemData        =   "testseiyuu.frx":0369
         Left            =   360
         List            =   "testseiyuu.frx":036B
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cbseiyuu 
         Height          =   315
         ItemData        =   "testseiyuu.frx":036D
         Left            =   360
         List            =   "testseiyuu.frx":036F
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   2895
      End
      Begin VB.Image imgseiyuu 
         BorderStyle     =   1  'Fixed Single
         Height          =   1695
         Left            =   4200
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Label lbltheme 
      BackColor       =   &H80000004&
      Caption         =   "&Thème :"
      Height          =   255
      Left            =   7080
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lbltitre 
      BackColor       =   &H80000004&
      Caption         =   "Pa&r :"
      Height          =   255
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblserie 
      BackColor       =   &H80000004&
      Caption         =   "&Anime :"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Seiyuu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'liste complète dans la combobox et liste par rapport à la sélection dans la listbox
'même gestion possible pour les persos
'avantages : la liste contient uniquement le résultat de la requête

'anime : une combobox contenant tous les titres
'seiyuu : une combobox contenant tous les seiyuu et une listbox contenant le résultat de la requête
'perso : idem seiyuu
'Initialisation de la liste lors de l'évènements keypress ou click du contrôle combobox
'Fusion du bouton rafraîchir avec le bouton fermer (base de données)
'Tous les contrôles visibles sauf peut-être avant l'ouverture de la BD ou après fermeture

'Par titre original ou titre français : à voir si possibilité de chercher dans les 2,
'ce qui supprimerait la combobox titre et élargirait la recherche
'Problème : voir aussi pour série, OAV ou film. (peut-être directement dans la BD)
'Hypothèse : une table Anime avec des tables "filles" Série, OAV et Film...
'exemple : Kenshin
'Anime (titre, titre origin, auteur)
'serie (nb episodes, année)
'film ('titre', persos en +,année)
'oav (titres, persos en +, nb oav, année)

Rem Modifier le curseur
Rem évènement dbl_click form3
Rem bon focus sur la listeiyuu

Option Explicit
Option Compare Text 'la casse n'est pas prise en compte
Public chemin As String 'chemin des images
Dim curseur As MousePointerConstants
Dim cn As New ADODB.Connection 'connexion ADODB
Dim rst As New ADODB.Recordset 'recordset
Dim fin As Boolean 'test de fermeture de la connexion
Dim doubl As Integer 'test de doublon
Dim autre As Boolean
Private Sub cmdouvrir_Click()
Dim req As String 'requête SQL

cn.Open "Toto", "chassajo", "ifurito" 'ouverture de la connexion

'remplissage de cbanime
req = "SELECT titre_anime FROM ANIME"
rst.Open req, cn, adOpenForwardOnly, adLockReadOnly
While Not rst.EOF
    cbanime.AddItem rst!titre_anime
    rst.MoveNext
Wend
rst.Close

'remplissage de cbseiyuu
req = "SELECT nom_seiyuu,prenom_seiyuu FROM SEIYUU"
rst.Open req, cn, adOpenForwardOnly, adLockReadOnly
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
req = "SELECT nom_perso,prenom_perso FROM PERSO"
rst.Open req, cn, adOpenForwardOnly, adLockReadOnly
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

fin = True 'test de fermeture de la connexion
cmdouvrir.Enabled = False
cmdfermer.Enabled = True
cbanime.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

cmdquitter_Click

End Sub

Private Sub photo_Click()

photo.Visible = False
photo.Top = 840

End Sub

Private Sub cbperso_Click()

cbperso_KeyPress (13)

End Sub

Private Sub cbperso_KeyPress(KeyAscii As Integer)
Dim i As Integer 'indice
Dim req As String 'requête SQL
Dim req2 As String 'requête SQL
Dim rst2 As New ADODB.Recordset 'recordset
Dim tablo() As String 'tableau contenant le nom et le prenom

If KeyAscii = 13 Then
    For i = 0 To cbperso.ListCount - 1
        If cbperso.Text = "" Then
            cbperso.ListIndex = -1
            Exit Sub
        ElseIf cbperso.List(i) Like cbperso.Text Then
            tablo = Split(cbperso.List(i), " ")
            'anime correspondant au perso
            If UBound(tablo) = 0 Then
                req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "'"
                req2 = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND prenom_perso = " & "'" & tablo(0) & "'"
                rst.Open req, cn, adOpenDynamic, adLockPessimistic
                rst2.Open req2, cn, adOpenDynamic, adLockPessimistic
                
                On Error Resume Next
                cbanime.Text = rst!titre_anime
                cbanime.Text = rst2!titre_anime
                On Error GoTo 0
                'rst.MoveFirst
                'rst2.MoveFirst
            Else
                req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = " & "'" & tablo(1) & "'"
                rst.Open req, cn, adOpenDynamic, adLockPessimistic
                cbanime.Text = rst!titre_anime
                'rst.Close
            End If
            
            doubl = 0
            Doublons.listdoublons.Clear
            If UBound(tablo) = 0 Then
                While Not rst.EOF
                    On Error GoTo prenom
                    If cbperso.List(i) = rst!nom_perso Then
                        doubl = doubl + 1
                        Doublons.listdoublons.AddItem rst!nom_perso
                    End If
                    rst.MoveNext
                Wend
prenom:         While Not rst2.EOF
                    If cbperso.List(i) = rst2!prenom_perso Then
                        doubl = doubl + 1
                        Doublons.listdoublons.AddItem rst2!prenom_perso
                    End If
                    rst2.MoveNext
                Wend
                rst2.Close
            Else
                While Not rst.EOF
                    If cbperso.List(i) = rst!nom_perso & " " & rst!prenom_perso Then
                        doubl = doubl + 1
                        Doublons.listdoublons.AddItem rst!nom_perso & " " & rst!prenom_perso
                    End If
                    rst.MoveNext
                Wend
            End If
            rst.Close
            
            'gestion des doublons perso
            If doubl > 1 Then
                Doublons.listdoublons.ListIndex = 0
                Doublons.Show vbModal
            End If
            
            cbanime_Click
            
            'Call focus_liste(3)
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucun perso correspondant", vbOKOnly + vbExclamation, "Perso"
    If cbperso.Text <> "" And Not fin Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Ouverture de la connexion"
End If

End Sub

Private Sub cbseiyuu_Click()

cbseiyuu_KeyPress (13)

End Sub

Private Sub cbanime_Click()

cbanime_KeyPress (13)

End Sub


Private Sub cbanime_KeyPress(KeyAscii As Integer)
Dim req As String 'requête SQL
Dim i As Integer 'indice
Dim j As Integer 'indice

If KeyAscii = 13 Then 'pression sur la touche "Entrée"
    'affichage des listes lors de la saisie d'un titre de série
    For i = 0 To cbanime.ListCount - 1
        If cbanime.Text = "" Then
            cbanime.ListIndex = -1
            Exit Sub
        ElseIf cbanime.List(i) Like cbanime.Text Then
            autre = False
            Call rempli_listseiyuu
            Call focus_liste(1)
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucun anime correspondant", vbOKOnly + vbExclamation, "Anime"
    If cbanime.Text <> "" And Not fin Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Ouverture de la connexion"
End If

End Sub

Private Sub cbtheme_Click()

If cbtheme.ListIndex = 0 Then
    change_theme (&HFFC0C0) 'bleu
ElseIf cbtheme.ListIndex = 1 Then
    change_theme (&HC0FFC0) 'vert
End If

End Sub

Private Sub cmdagrandir_Click(Index As Integer)
'affichage de l'aperçu d'une photo
'géré avec une autre form ou avec une picturebox
'index = photo sélectionnée

If Index = 0 Then
    If imgseiyuu.Picture = 0 Then
        MsgBox "Aucune photo sélectionnée", vbOKOnly + vbCritical, "Erreur photo"
        Exit Sub
    Else
        Visuel.Picture = imgseiyuu.Picture
        Visuel.Show
    End If
ElseIf Index = 1 Then
    If imgperso.Picture = 0 Then
        MsgBox "Aucune photo sélectionnée", vbOKOnly + vbCritical, "Erreur photo"
    Else
        Visuel.Picture = imgperso.Picture
        Visuel.Show
    End If
Else
    MsgBox "Impossible d'afficher l'photo sélectionnée", vbCritical + vbOKOnly, "Erreur photo"
End If

'affichage de l'aperçu d'une photo
'gérer avec une picturebox
'index = photo sélectionnée
'If Index = 0 Then
'    photo.Visible = True
'    photo = imgseiyuu.Picture
'ElseIf Index = 1 Then
'    photo.Visible = True
'    photo.Top = 3480
'    photo = imgperso.Picture
'Else
'    MsgBox "Impossible d'afficher l'photo sélectionnée", vbCritical + vbOKOnly, "Erreur photo"
'End If

End Sub

Private Sub cmdagrandir_LostFocus(Index As Integer)

photo_Click

End Sub

Private Sub cmdquitter_Click()

'fermeture de la connexion
If fin Then
    cn.Close
End If

'déchargement des feuilles
Unload Doublons
Unload Visuel

End 'fermeture du programme

End Sub

Private Sub Form_Load()

fin = False 'test de fermeture de la connexion
chemin = "C:\WINNT\Profiles\Administrateur\Bureau\Projet\" 'chemin des images
curseur = vbArrowQuestion
frmseiyuu.MousePointer = curseur
cbtitre.ListIndex = 0
cbtheme.ListIndex = 0 'initialisation du thème
listseiyuu.Visible = False
listperso.Visible = False
photo.Visible = False
cmdagrandir(0).Enabled = False
cmdagrandir(1).Enabled = False
cmdfermer.Enabled = False

End Sub

Private Sub listperso_Click()
'image par rapport au perso sélectionné

Dim req As String 'requête SQL
Dim tablo() As String 'tableau contenant le nom et le prenom
Dim j As Integer 'indice
Dim req2 As String 'requête SQL
Dim rst2 As New ADODB.Recordset 'recordset

tablo = Split(listperso.List(listperso.ListIndex), " ")
If UBound(tablo) = 0 Then
    req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "'"
    req2 = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND prenom_perso = " & "'" & tablo(0) & "'"
    rst.Open req, cn, adOpenDynamic, adLockPessimistic
    rst2.Open req2, cn, adOpenDynamic, adLockPessimistic
                
    On Error Resume Next
    cbanime.Text = rst!titre_anime
    imgperso.Picture = LoadPicture(chemin & rst!nom_image)
    imgperso.Tag = rst!nom_image
    cbanime.Text = rst2!titre_anime
    imgperso.Picture = LoadPicture(chemin & rst2!nom_image)
    imgperso.Tag = rst2!nom_image
    On Error GoTo 0
    rst2.Close
Else
    req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = " & "'" & tablo(1) & "'"
    rst.Open req, cn, adOpenDynamic, adLockPessimistic
    cbanime.Text = rst!titre_anime
    'imgperso.Picture = LoadPicture(chemin & rst!nom_image)
    imgperso.Tag = rst!nom_image
End If
rst.Close
Call rempli_listseiyuu
imgperso.ToolTipText = "image de " & listperso.List(listperso.ListIndex)
cmdagrandir(1).Enabled = True
cbperso.Text = listperso.List(listperso.ListIndex)
autre = True
Call focus_liste(3)
Exit Sub

End Sub

Private Sub listseiyuu_Click()
'list des persos par rapport au seiyuu
'photo par rapport au seiyuu sélectionné

Dim req As String 'requête SQL
Dim tablo() As String 'tableau contenant le nom et le prenom
Dim j As Integer 'indice
Dim i As Integer 'indice

If Not autre Then
    tablo = Split(listseiyuu.List(listseiyuu.ListIndex), " ")
    req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,photo_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_seiyuu = " & "'" & tablo(0) & "' AND prenom_seiyuu = " & "'" & tablo(1) & "'"
    rst.Open req, cn, adOpenDynamic, adLockPessimistic

    'imgseiyuu.Picture = LoadPicture(chemin & rst!photo_seiyuu)
    'remplissage de la liste des persos
    listperso.Clear
    listperso.Visible = True
    j = 0
    While Not rst.EOF
        If Not IsNull(rst!nom_perso) Then
            If Not IsNull(rst!prenom_perso) Then
                listperso.AddItem rst!nom_perso & " " & rst!prenom_perso
                'Doublons.listdoublons.AddItem rst!nom_perso & " " & rst!prenom_perso
            Else
                listperso.AddItem rst!nom_perso
                'Doublons.listdoublons.AddItem rst!nom_perso
            End If
        ElseIf IsNull(rst!nom_perso) And Not IsNull(rst!prenom_perso) Then
            listperso.AddItem rst!prenom_perso
            'Doublons.listdoublons.AddItem rst!prenom_perso
        End If
        rst.MoveNext
        Call supp_doublons(j, listperso)
    Wend
    rst.Close

    'gestion des doublons perso
    'If cbperso.Text Like listperso.List(listperso.ListCount - 1) And doubl > 1 Then
    '    Doublons.listdoublons.ListIndex = 0
    '    Doublons.Show vb modal
    'End If
End If

Call focus_liste(2)
imgseiyuu.ToolTipText = "photo de " & listseiyuu.List(listseiyuu.ListIndex)
cmdagrandir(0).Enabled = True
cbseiyuu.Text = listseiyuu.List(listseiyuu.ListIndex)

End Sub

Private Sub cbseiyuu_KeyPress(KeyAscii As Integer)
Dim i As Integer 'indice
Dim req As String 'requête SQL
Dim j As Integer 'indice
Dim tablo() As String 'tableau contenant le nom et le prenom

If KeyAscii = 13 Then
    'affichage des listes lors de la saisie du nom d'un seiyuu
    For i = 0 To cbseiyuu.ListCount - 1
        If cbseiyuu.Text = "" Then
            cbseiyuu.ListIndex = -1
            Exit Sub
        ElseIf cbseiyuu.List(i) Like cbseiyuu.Text Then
            tablo = Split(cbseiyuu.List(i), " ")
            req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_seiyuu = " & "'" & tablo(0) & "' AND prenom_seiyuu = " & "'" & tablo(1) & "'"
            rst.Open req, cn, adOpenDynamic, adLockPessimistic
            cbanime.Text = rst!titre_anime
            rst.Close
            cbanime_Click
            'Call focus_liste(2)
            Exit Sub
        End If
    Next i
    If i <> 0 Then MsgBox "Aucun(e) seiyuu correspondant(e)", vbOKOnly + vbExclamation, "Seiyuu"
    If cbseiyuu.Text <> "" And Not fin Then MsgBox "Veuillez ouvrir la connexion", vbExclamation + vbOKOnly, "Ouverture de la connexion"
End If

End Sub

Private Sub cmdfermer_click()
'initialisation des contrôles

listseiyuu.Visible = False
listperso.Visible = False
imgseiyuu.Picture = LoadPicture()
imgperso.Picture = LoadPicture()
cbperso.Text = ""
cbseiyuu.Text = ""
cbanime.Text = ""
cbanime.Clear
cbseiyuu.Clear
cbperso.Clear
listseiyuu.Clear
listperso.Clear
photo.Top = 840
photo.Picture = LoadPicture()
photo.Visible = False
cmdagrandir(0).Enabled = False
cmdagrandir(1).Enabled = False
cmdouvrir.Enabled = True
cmdfermer.Enabled = False
cbanime.SetFocus

'fermeture de la connexion
If fin Then
    cn.Close
End If

fin = False 'test de fermeture de la connexion

End Sub

Private Sub change_theme(couleur As String)
'changement du thème de l'application

Seiyuu.BackColor = couleur
frmseiyuu.BackColor = couleur
frmperso.BackColor = couleur
lbltitre.BackColor = couleur
lblserie.BackColor = couleur
lbltheme.BackColor = couleur
Doublons.BackColor = couleur
Doublons.lbldoublons.BackColor = couleur
Doublons.listdoublons.BackColor = couleur
Visuel.BackColor = couleur
photo.BackColor = couleur

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
Private Sub rempli_listseiyuu()
Dim req As String 'requête SQL
Dim j As Integer 'indice

req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND titre_anime = " & "'" & cbanime.Text & "'"
rst.Open req, cn, adOpenDynamic, adLockPessimistic
            
'remplissage de la liste des seiyuu
listseiyuu.Clear
listseiyuu.Visible = True
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
    Call supp_doublons(j, listseiyuu)
Wend
rst.Close

End Sub
Private Sub focus_liste(num_liste As Integer)
'gestion des focus sur les listes

Dim i As Integer 'indice
Dim req As String 'requête SQL
Dim req2 As String 'requête SQL
Dim rst2 As New ADODB.Recordset 'recordset
Dim tablo() As String 'tableau contenant le nom et le prenom

Select Case num_liste
    Case 1
        listseiyuu.ListIndex = 0
    Case 2
        tablo = Split(listseiyuu.List(listseiyuu.ListIndex), " ")
        'un anime correspondant au seiyuu
        req = "SELECT titre_anime,nom_seiyuu,prenom_seiyuu,photo_seiyuu,nom_perso,prenom_perso FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGE WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_seiyuu = " & "'" & tablo(0) & "' AND prenom_seiyuu = " & "'" & tablo(1) & "' AND titre_anime = " & "'" & cbanime.Text & "'"
        rst.Open req, cn, adOpenDynamic, adLockPessimistic
        For i = 0 To listperso.ListCount
            If (listperso.List(i) = rst!nom_perso & " " & rst!prenom_perso) Or (listperso.List(i) = rst!nom_perso) Or (listperso.List(i) = rst!prenom_perso) Then
                rst.Close
                listperso.ListIndex = i 'focus sur le perso de l'anime doublé par le seiyuu
                Exit For
            End If
        Next i
    Case 3
        tablo = Split(listperso.List(listperso.ListIndex))
        'seiyuu correspondant au perso
        If UBound(tablo) = 0 Then
            req = "SELECT nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM PERSO,SEIYUU,DOUBLER,IMAGE WHERE PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "' AND nom_image = " & "'" & imgperso.Tag & "'"
            req2 = "SELECT nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM PERSO,SEIYUU,DOUBLER,IMAGE WHERE PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND prenom_perso = " & "'" & tablo(0) & "' AND nom_image = " & "'" & imgperso.Tag & "'"
            rst.Open req, cn, adOpenDynamic, adLockPessimistic
            rst2.Open req2, cn, adOpenDynamic, adLockPessimistic
            
            For i = 0 To listseiyuu.ListCount
                On Error GoTo prenom
                If (listseiyuu.List(i) = rst!nom_seiyuu & " " & rst!prenom_seiyuu) Then
                    rst.Close
                    rst2.Close
                    listseiyuu.ListIndex = i 'focus sur le seiyuu correspondant au perso
                    Exit For
                End If
            Next i
prenom:     For i = 0 To listseiyuu.ListCount
                If (listseiyuu.List(i) = rst2!nom_seiyuu & " " & rst2!prenom_seiyuu) And (rst2!nom_image = imgperso.Tag) Then
                    rst.Close
                    rst2.Close
                    listseiyuu.ListIndex = i 'focus sur le seiyuu correspondant au perso
                    Exit For
                End If
            Next i
        Else
            req = "SELECT nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,nom_image FROM PERSO,SEIYUU,DOUBLER,IMAGE WHERE PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_image = IMAGE.num_image AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = " & "'" & tablo(1) & "'"
            rst.Open req, cn, adOpenDynamic, adLockPessimistic
            For i = 0 To listseiyuu.ListCount
                If (listseiyuu.List(i) = rst!nom_seiyuu & " " & rst!prenom_seiyuu) Then
                    rst.Close
                    listseiyuu.ListIndex = i 'focus sur le seiyuu correspondant au perso
                    Exit For
                End If
            Next i
        End If
End Select

End Sub
