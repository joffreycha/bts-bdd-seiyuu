VERSION 5.00
Begin VB.Form Ajouter 
   Caption         =   "Ajouter"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optchoix 
      Caption         =   "&Série"
      Height          =   255
      Index           =   2
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optchoix 
      Caption         =   "&Doubleur"
      Height          =   255
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optchoix 
      Caption         =   "&Personnage"
      Height          =   255
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdannuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Annuler la création"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtnomperso 
      Height          =   285
      Left            =   6000
      TabIndex        =   8
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtprenomperso 
      Height          =   285
      Left            =   6000
      TabIndex        =   10
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtanime 
      Height          =   285
      Left            =   1440
      TabIndex        =   12
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtprenomseiyuu 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtnomseiyuu 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblanime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Série dans laquelle le personnage apparait :"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   3090
   End
   Begin VB.Label lblprenomseiyuu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prénom du doubleur :"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1515
   End
   Begin VB.Label lblnomseiyuu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du doubleur :"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1305
   End
   Begin VB.Label lblprenomperso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prénom du personnage :"
      Height          =   195
      Left            =   4200
      TabIndex        =   9
      Top             =   1200
      UseMnemonic     =   0   'False
      Width           =   1740
   End
   Begin VB.Label lblnomperso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nom du personnage :"
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   600
      UseMnemonic     =   0   'False
      Width           =   1530
   End
End
Attribute VB_Name = "Ajouter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force la déclaration des variables

Dim rst As New ADODB.Recordset
Private Function increment_num(table As String) As Integer
'incrémentation du numéro identifiant d'une table entrée en paramètre

Dim rstnum As New ADODB.Recordset

rstnum.Open "SELECT MAX(num_" & table & ") As nummaxi FROM " & table, Seiyuu.cn, adOpenKeyset, adLockPessimistic
    increment_num = rstnum!nummaxi + 1
rstnum.Close

End Function

Private Sub vide_feuille()
'initialise les contrôles

txtnomseiyuu.Text = ""
txtprenomseiyuu.Text = ""
txtnomperso.Text = ""
txtprenomperso.Text = ""
txtanime.Text = ""

End Sub

Private Sub cmdannuler_Click()

Unload Ajouter

End Sub

Private Sub cmdok_Click()
'ajout ou modification d'un personnage, d'un seiyuu ou d'un anime

If optchoix(0).Value = True Then 'création ou modification d'un personnage
    'refus si certains champs sont vides
    If txtnomperso.Text <> "" Or txtprenomperso.Text <> "" Then
        'création d'un personnage
        If Seiyuu.modif = 0 Then
            'demande de confirmation de la création
            If MsgBox("Êtes-vous sûr de vouloir créer ce personnage ?", vbQuestion + vbYesNo, "Création personnage") = vbYes Then
                'déclaration des recordset
                Dim rstperso As New ADODB.Recordset
                Dim rstimage As New ADODB.Recordset
                Dim rstseiyuu As New ADODB.Recordset
                Dim rstanime As New ADODB.Recordset
                Dim rstdoubler As New ADODB.Recordset
                Dim rstapp As New ADODB.Recordset
            
                'création du tuple dans la table PERSO
                rstperso.Open "SELECT * FROM PERSO", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
                    rstperso.AddNew
                    rstperso!num_perso = increment_num("perso") 'incrémentation du numéro identifiant
                    If txtnomperso.Text <> "" Then rstperso!nom_perso = txtnomperso.Text
                    If txtprenomperso.Text <> "" Then rstperso!prenom_perso = txtprenomperso.Text
                
                'création du tuple dans la table IMAGES
                rstimage.Open "SELECT * FROM IMAGES", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
                    rstimage.AddNew
                    rstimage!num_images = increment_num("images") 'incrémentation du numéro identifiant
                    
                    'nom de l'image
                    If txtnomperso.Text <> "" And txtprenomperso.Text <> "" Then
                        rstimage!nom_images = txtnomperso.Text & "_" & txtprenomperso.Text & ".jpg"
                    ElseIf txtnomperso.Text <> "" Then
                        rstimage!nom_images = txtnomperso.Text & ".jpg"
                    ElseIf txtprenomperso.Text <> "" Then
                        rstimage!nom_images = txtprenomperso.Text & ".jpg"
                    End If

                'récupération des identifiants
                rstseiyuu.Open "SELECT * FROM SEIYUU WHERE nom_seiyuu = '" & txtnomseiyuu.Text & "' AND prenom_seiyuu = '" & txtprenomseiyuu.Text & "'", Seiyuu.cn, adOpenKeyset, adLockOptimistic
        
                rstanime.Open "SELECT * FROM ANIME WHERE titre_original = '" & txtanime.Text & "'", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
        
                'création du tuple dans la table DOUBLER
                rstdoubler.Open "SELECT * FROM DOUBLER", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
                    rstdoubler.AddNew
                    rstdoubler!num_perso = rstperso!num_perso
                    On Error GoTo erreurseiyuu 'identifiant non trouvé
                    rstdoubler!num_seiyuu = rstseiyuu!num_seiyuu
                    rstdoubler!num_images = rstimage!num_images

                'création du tuple dans la table APPARAITRE
                rstapp.Open "SELECT * FROM APPARAITRE", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
                    rstapp.AddNew
                    rstapp!num_perso = rstperso!num_perso
                    On Error GoTo erreuranime 'identifiant non trouvé
                    rstapp!num_anime = rstanime!num_anime
            
                'mise à jour
                rstperso.Update
                rstimage.Update
                rstdoubler.Update
                rstapp.Update
            
                'fermeture des recordset
                rstperso.Close
                rstimage.Close
                rstseiyuu.Close
                rstanime.Close
                rstdoubler.Close
                rstapp.Close
        
            Else
                Exit Sub
            End If
        Else 'modification d'un personnage
        
            'demande de confirmation de la modification
            If MsgBox("Êtes-vous sûr de vouloir modifier ce personnage ?", vbQuestion + vbYesNo, "Modification personnage") = vbYes Then
                rst.Open "SELECT num_seiyuu, nom_seiyuu, prenom_seiyuu FROM SEIYUU WHERE nom_seiyuu = '" & txtnomseiyuu.Text & "' AND prenom_seiyuu = '" & txtprenomseiyuu.Text & "'", Seiyuu.cn, adOpenKeyset, adLockPessimistic
            
                'récupération des identifiants
                On Error GoTo erreurseiyuu 'identifiant non trouvé
                Seiyuu.idseiyuu = rst!num_seiyuu
                rst.Close
            
                rst.Open "SELECT num_anime, titre_original FROM ANIME WHERE titre_original = '" & txtanime.Text & "'", Seiyuu.cn, adOpenKeyset, adLockPessimistic
                On Error GoTo erreuranime 'identifiant non trouvé
                Seiyuu.idanime = rst!num_anime
                On Error GoTo 0
                rst.Close
            
                'exécution de la procédure stockée
                'paramètres : num_perso,nom_perso,prenom_perso,num_seiyuu,num_anime
                rst.Open "modifier_perso '" & Seiyuu.idperso & "','" & txtnomperso.Text & "','" & txtprenomperso.Text & "','" & Seiyuu.idseiyuu & "','" & Seiyuu.idanime & "'"
            
                'modification du nom de l'image
                rst.Open "SELECT * FROM IMAGES WHERE num_images = '" & Seiyuu.idimages & "'"
                If txtnomperso.Text <> "" And txtprenomperso.Text <> "" Then
                    rst!nom_images = txtnomperso.Text & "_" & txtprenomperso.Text & ".jpg"
                ElseIf txtnomperso.Text <> "" Then
                    rst!nom_images = txtnomperso.Text & ".jpg"
                ElseIf txtprenomperso.Text <> "" Then
                    rst!nom_images = txtprenomperso.Text & ".jpg"
                End If
                rst.Update
            Else
                Exit Sub
            End If
        End If
    Else
        MsgBox "Vous devez saisir au moins le nom ou le prénom du personnage !", vbOKOnly + vbExclamation, "Champs vides"
        Exit Sub
    End If
    
ElseIf optchoix(1).Value = True Then 'création ou modification d'un seiyuu
    'refus si certains champs sont vides
    If txtnomseiyuu.Text <> "" And txtprenomseiyuu.Text <> "" Then
        'création d'un seiyuu
        If Seiyuu.modif = 0 Then
            rst.Open "SELECT * FROM SEIYUU", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            
                'création du tuple dans la table SEIYUU
                rst.AddNew
                rst!num_seiyuu = increment_num("seiyuu")
                rst!nom_seiyuu = txtnomseiyuu.Text
                rst!prenom_seiyuu = txtprenomseiyuu.Text
                rst!photo_seiyuu = txtnomseiyuu.Text & "_" & txtprenomseiyuu.Text & ".jpg"
        
                'demande de confirmation de la création
                If MsgBox("Êtes-vous sûr de vouloir créer ce doubleur ?", vbQuestion + vbYesNo, "Création personnage") = vbYes Then
                    rst.Update
                Else
                    rst.CancelUpdate
                End If
            rst.Close
        Else 'modification d'un seiyuu
        
            'demande de confirmation de la modification
            If MsgBox("Êtes-vous sûr de vouloir modifier ce doubleur ?", vbQuestion + vbYesNo, "Modification doubleur") = vbYes Then
            
                'exécution de la procédure stockée
                'paramètres : num_seiyuu,nom_seiyuu,prenom_seiyuu,photo_seiyuu
                rst.Open "modifier_seiyuu '" & Seiyuu.idseiyuu & "','" & txtnomseiyuu.Text & "','" & txtprenomseiyuu.Text & "','" & txtnomseiyuu.Text & "_" & txtprenomseiyuu.Text & ".jpg"
            End If
        End If
    Else
        MsgBox "Vous devez obligatoirement saisir le nom et le prénom du doubleur !", vbOKOnly + vbExclamation, "Champs vides"
        Exit Sub
    End If
            
ElseIf optchoix(2).Value = True Then 'création ou modification d'un anime
    'refus si certains champs sont vides
    If txtnomseiyuu.Text <> "" Then
        'création d'un anime
        If Seiyuu.modif = 0 Then
            rst.Open "SELECT * FROM ANIME", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
                'création du tuple dans la table ANIME
                rst.AddNew
                rst!num_anime = increment_num("anime") 'incrémentation du numéro identifiant
                rst!titre_original = txtnomseiyuu.Text
                If txtprenomseiyuu.Text <> "" Then rst!titre_version_fr = txtprenomseiyuu.Text
                If txtnomperso.Text <> "" Then rst!generique_deb = txtnomperso.Text
                If txtprenomperso.Text <> "" Then rst!generique_fin = txtprenomperso.Text
            
                'demande de confirmation de la création
                If MsgBox("Êtes-vous sûr de vouloir créer cette série ?", vbQuestion + vbYesNo, "Création personnage") = vbYes Then
                    rst.Update
                Else
                    rst.CancelUpdate
                End If
            rst.Close
        Else 'modification d'un anime
            'demande de confirmation de la modification
            If MsgBox("Êtes-vous sûr de vouloir modifier cette série ?", vbQuestion + vbYesNo, "Modification série") = vbYes Then
                
                'exécution de la procédure stockée
                'paramètres : num_anime,titre_original,titre_version_fr,generique_deb,generique_fin
                rst.Open "modifier_anime '" & Seiyuu.idanime & "','" & txtnomseiyuu.Text & "','" & txtprenomseiyuu.Text & "','" & txtnomperso.Text & "','" & txtprenomperso.Text & "'"
            End If
        End If
    Else
        MsgBox "Vous devez obligatoirement saisir le titre original de la série !", vbOKOnly + vbExclamation, "Champs vides"
        Exit Sub
    End If
End If

'Mise à jour des listes
Seiyuu.Mfermer_Click
Seiyuu.Mouvrir_Click

Unload Ajouter 'déchargement de la feuille

Exit Sub

erreurseiyuu:
MsgBox "Il n'existe aucun doubleur de ce nom, veuillez corriger la saisie ou l'enregistrer en tant que nouveau doubleur.", vbOKOnly + vbCritical, "Erreur doubleur"
If Seiyuu.modif = 0 Then
    
    'annulation de la mise à jour
    rstperso.CancelUpdate
    rstimage.CancelUpdate
    rstdoubler.CancelUpdate
            
    'fermeture des recordset
    rstperso.Close
    rstimage.Close
    rstseiyuu.Close
    rstanime.Close
    rstdoubler.Close
Else
    rst.Close
End If

Exit Sub

erreuranime:
MsgBox "Il n'existe aucune série de ce nom, veuillez corriger la saisie ou l'enregistrer en tant que nouvelle série.", vbOKOnly + vbCritical, "Erreur série"
If Seiyuu.modif = 0 Then
    'annulation de la mise à jour
    rstperso.CancelUpdate
    rstimage.CancelUpdate
    rstdoubler.CancelUpdate
    rstapp.CancelUpdate
            
    'fermeture des recordset
    rstperso.Close
    rstimage.Close
    rstseiyuu.Close
    rstanime.Close
    rstdoubler.Close
    rstapp.Close
Else
    rst.Close
End If

End Sub

Private Sub Form_Load()

Select Case Seiyuu.modif
    Case 0 'ajout
        optchoix(0).Value = True

    Case 1 'modification de personnage
        optchoix(0).Value = True
        
    Case 2 'modification de seiyuu
        optchoix(1).Value = True
        
    Case 3 'une modification d'anime
        optchoix(2).Value = True
    
End Select
Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name) 'changement du thème

End Sub

Private Sub rempli_feuille(modification As Integer)
'complète la form Ajouter avec les informations à mettre à jour

Dim tablo() As String 'tableau contenant les nom et prénom de personnage

Select Case modification
    Case 0 'ajout
        Ajouter.Caption = "Ajouter"
        
    Case 1 'modification de personnage
        Dim tablo2() As String 'tableau contenant les nom et prénom de seiyuu

        Call vide_feuille 'initialisation des contrôles
        
        Ajouter.Caption = "Modifier"
        
        tablo = Split(Seiyuu.listperso.List(Seiyuu.listperso.ListIndex))
        tablo2 = Split(Seiyuu.listseiyuu.List(Seiyuu.listseiyuu.ListIndex))
        
        If UBound(tablo) = 0 Then 'seulement un nom ou un prénom
            req = "SELECT SEIYUU.num_seiyuu,ANIME.num_anime,PERSO.num_perso,titre_original,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,IMAGES.num_images FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = '" & tablo2(0) & "' AND prenom_seiyuu = '" & tablo2(1) & "' AND (nom_perso = " & "'" & tablo(0) & "' OR prenom_perso = '" & tablo(0) & "')"
        Else 'nom et prénom
            req = "SELECT SEIYUU.num_seiyuu,ANIME.num_anime,PERSO.num_perso,titre_original,nom_seiyuu,prenom_seiyuu,nom_perso,prenom_perso,IMAGES.num_images FROM ANIME,SEIYUU,DOUBLER,PERSO,APPARAITRE,IMAGES WHERE ANIME.num_anime = APPARAITRE.num_anime AND APPARAITRE.num_perso = PERSO.num_perso AND PERSO.num_perso = DOUBLER.num_perso AND DOUBLER.num_seiyuu = SEIYUU.num_seiyuu AND DOUBLER.num_images = IMAGES.num_images AND nom_seiyuu = '" & tablo2(0) & "' AND prenom_seiyuu = '" & tablo2(1) & "' AND nom_perso = " & "'" & tablo(0) & "' AND prenom_perso = '" & tablo(1) & "'"
        End If
        
        rst.Open req, Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
        
        'remplissage des contrôles
        txtnomseiyuu.Text = rst!nom_seiyuu
        txtprenomseiyuu.Text = rst!prenom_seiyuu
        On Error Resume Next
        txtnomperso.Text = rst!nom_perso
        txtprenomperso.Text = rst!prenom_perso
        On Error GoTo 0
        txtanime.Text = rst!titre_original
        
        rst.Close
        
    Case 2 'modification de seiyuu
        Call vide_feuille 'initialisation des contrôles
        
        Ajouter.Caption = "Modifier"
        tablo() = Split(Seiyuu.listseiyuu.List(Seiyuu.listseiyuu.ListIndex))
        rst.Open "SELECT * FROM SEIYUU WHERE nom_seiyuu = '" & tablo(0) & "' AND prenom_seiyuu = '" & tablo(1) & "'", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            txtnomseiyuu.Text = rst!nom_seiyuu
            txtprenomseiyuu.Text = rst!prenom_seiyuu
        rst.Close
        
    Case 3 'modification d'anime
        Call vide_feuille 'initialisation des contrôles
        
        Ajouter.Caption = "Modifier"
        rst.Open "SELECT * FROM ANIME WHERE (titre_original = '" & Seiyuu.cbanime.Text & "' OR titre_version_fr = '" & Seiyuu.cbanime.Text & "')", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            txtnomseiyuu.Text = rst!titre_original
            If Not IsNull(rst!titre_version_fr) Then txtprenomseiyuu.Text = rst!titre_version_fr
            If Not IsNull(rst!generique_deb) Then txtnomperso.Text = rst!generique_deb
            If Not IsNull(rst!generique_fin) Then txtprenomperso.Text = rst!generique_fin
        rst.Close
    
End Select

End Sub

Private Sub optchoix_Click(Index As Integer)
'changement de l'interface de la form

Select Case Index
    Case 0 'personnage
        lblnomseiyuu.Caption = "Nom du doubleur :"
        lblprenomseiyuu.Caption = "Prénom du doubleur :"
        With lblnomperso
            .Visible = True
            .Caption = "Nom du personnage :"
        End With
        txtnomperso.Visible = True
        With lblprenomperso
            .Visible = True
            .Caption = "Prénom du personnage :"
        End With
        txtprenomperso.Visible = True
        With lblanime
            .Visible = True
            .Caption = "Série dans laquelle le personnage apparait : (titre original)"
        End With
        txtanime.Visible = True
        
    Case 1 'seiyuu
        lblnomseiyuu.Caption = "Nom du doubleur :"
        lblprenomseiyuu.Caption = "Prénom du doubleur :"
        
        lblnomperso.Visible = False
        lblprenomperso.Visible = False
        txtnomperso.Visible = False
        txtprenomperso.Visible = False
        lblanime.Visible = False
        txtanime.Visible = False
    
    Case 2 'anime
        lblnomseiyuu.Caption = "Titre original de la série :"
        lblprenomseiyuu.Caption = "Titre version française :"
        With lblnomperso
            .Visible = True
            .Caption = "Générique de début :"
        End With
        txtnomperso.Visible = True
        With lblprenomperso
            .Visible = True
            .Caption = "Générique de fin :"
        End With
        txtprenomperso.Visible = True
        lblanime.Visible = False
        txtanime.Visible = False
End Select

'complète la form
If Seiyuu.modif <> 0 Then Call rempli_feuille(Index + 1)

End Sub
