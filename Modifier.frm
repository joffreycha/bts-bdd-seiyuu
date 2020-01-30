VERSION 5.00
Begin VB.Form Modifier 
   Caption         =   "Ajouter"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   8160
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optchoix 
      Caption         =   "&Série"
      Height          =   255
      Index           =   2
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optchoix 
      Caption         =   "&Doubleur"
      Height          =   255
      Index           =   1
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   0
      Width           =   1215
   End
   Begin VB.OptionButton optchoix 
      Caption         =   "&Personnage"
      Height          =   255
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdannuler 
      Cancel          =   -1  'True
      Caption         =   "&Annuler"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
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
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtnomperso 
      Height          =   285
      Left            =   6000
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox txtprenomperso 
      Height          =   285
      Left            =   6000
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtanime 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtprenomseiyuu 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtnomseiyuu 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
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
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   8
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
Attribute VB_Name = "Modifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As New ADODB.Recordset
Private Sub cmdannuler_Click()

Unload Ajouter

End Sub

Private Sub cmdok_Click()

If optchoix(0).Value = True Then
    'création d'un nouveau personnage
    
    If MsgBox("Êtes-vous sûr de vouloir créer ce personnage ?", vbQuestion + vbYesNo, "Création personnage") = vbYes Then
        Dim rstimage As New ADODB.Recordset
        Dim rstperso As New ADODB.Recordset
        Dim rstseiyuu As New ADODB.Recordset
        Dim rstanime As New ADODB.Recordset
        
        'création du tuple dans la table PERSO
        rst.Open "SELECT * FROM PERSO", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            rst.AddNew
            rst!nom_perso = txtnomperso.Text
            rst!prenom_perso = txtprenomperso.Text
            rst.Update
        rst.Close
    
        rstimage.Open "SELECT * FROM IMAGES", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            rstimage.AddNew "nom_image", txtnomperso.Text & "_" & txtprenomperso.Text & ".jpg"
            rstimage.Update
        rstimage.Close
        
        rstimage.Open "SELECT * FROM IMAGES WHERE nom_image = '" & txtnomperso.Text & "_" & txtprenomperso.Text & ".jpg'", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
        
        rstperso.Open "SELECT * FROM PERSO WHERE nom_perso = '" & txtnomperso.Text & "' AND prenom_perso = '" & txtprenomperso.Text & "'", Seiyuu.cn, adOpenKeyset, adLockOptimistic
    
        rstseiyuu.Open "SELECT * FROM SEIYUU WHERE nom_seiyuu = '" & txtnomseiyuu.Text & "' AND prenom_seiyuu = '" & txtprenomseiyuu.Text & "'", Seiyuu.cn, adOpenKeyset, adLockOptimistic
        
        rstanime.Open "SELECT * FROM ANIME WHERE titre_original = '" & txtanime.Text & "'", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
        
        'création du tuple dans la table DOUBLER
        rst.Open "SELECT * FROM DOUBLER", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            rst.AddNew
            rst!num_perso = rstperso!num_perso
            rst!num_seiyuu = rstseiyuu!num_seiyuu
            rst!num_image = rstimage!num_image
            rst.Update
        rst.Close
        
        'création du tuple dans la table APPARAITRE
        rst.Open "SELECT * FROM APPARAITRE", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
            rst.AddNew
            rst!num_perso = rstperso!num_perso
            rst!num_anime = rstanime!num_anime
            rst.Update
        rst.Close
        rstimage.Close
        rstperso.Close
        rstseiyuu.Close
        rstanime.Close
    End If

ElseIf optchoix(1).Value = True Then
    'création d'un nouveau doubleur
    
    rst.Open "SELECT * FROM SEIYUU", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
        'ajout du nouveau tuple
        rst.AddNew
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
    
ElseIf optchoix(2).Value = True Then
    'création d'une nouvelle série
    
    rst.Open "SELECT * FROM ANIME", Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
    'ajout du nouveau tuple
        rst.AddNew
        rst!titre_original = txtnomseiyuu.Text
        rst!titre_version_fr = txtprenomseiyuu.Text
        rst!generique_deb = txtnomperso.Text
        rst!generique_fin = txtprenomperso.Text
        
        'demande de confirmation de la création
        If MsgBox("Êtes-vous sûr de vouloir créer cette série ?", vbQuestion + vbYesNo, "Création personnage") = vbYes Then
            rst.Update
        Else
            rst.CancelUpdate
        End If
    rst.Close
End If

'Mise à jour des listes
Seiyuu.Mfermer_Click
Seiyuu.Mouvrir_Click

Unload Ajouter

End Sub

Private Sub Form_Load()

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name)
optchoix(0).Value = True

End Sub

Private Sub optchoix_Click(Index As Integer)

Select Case Index
    Case 0
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
            .Caption = "Série dans laquelle le personnage apparait :"
        End With
        txtanime.Visible = True
        
    Case 1
        lblnomseiyuu.Caption = "Nom du doubleur :"
        lblprenomseiyuu.Caption = "Prénom du doubleur :"
        
        lblnomperso.Visible = False
        lblprenomperso.Visible = False
        txtnomperso.Visible = False
        txtprenomperso.Visible = False
        lblanime.Visible = False
        txtanime.Visible = False
    
    Case 2
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

End Sub
