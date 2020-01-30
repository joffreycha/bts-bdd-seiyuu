VERSION 5.00
Begin VB.Form Explorateur 
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdouvrir 
      Caption         =   "&Ouvrir"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ComboBox cbtype 
      Height          =   315
      ItemData        =   "Explorateur.frx":0000
      Left            =   1320
      List            =   "Explorateur.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton cmdannuler 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdenregistrer 
      Caption         =   "&Enregistrer"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox txtnomfic 
      Height          =   285
      Left            =   1320
      TabIndex        =   5
      Top             =   3000
      Width           =   3375
   End
   Begin VB.FileListBox fichiers 
      Height          =   2040
      Left            =   3480
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.DriveListBox disc 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.DirListBox dossiers 
      Height          =   2115
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lbltype 
      BackStyle       =   0  'Transparent
      Caption         =   "&Type :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label lblnomfic 
      BackStyle       =   0  'Transparent
      Caption         =   "&Nom du fichier :"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lblenregistrer 
      BackStyle       =   0  'Transparent
      Caption         =   "Enregistrer &dans :"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Explorateur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force la déclaration des variables
Option Compare Text 'ne prend pas en compte la casse

Private Sub cbtype_Click()
'change le type de fichier à ouvrir

fichiers.Pattern = cbtype.Text
txtnomfic.Text = cbtype.Text
txtnomfic.SelLength = Len(txtnomfic.Text)
fichiers.Path = dossiers.Path

End Sub

Private Function ajoute_audio(typefic As String) As Boolean
'ajoute un fichier audio à la playslist
'typefic : type du fichier à ajouter

Dim i As Integer 'indice

If fichiers.ListIndex <> -1 Then
    Lecteur.listplaylist.AddItem fichiers.List(fichiers.ListIndex)
Else
    'ajout de l'extension si elle est absente
    If Right(txtnomfic.Text, 4) = typefic Then
        Lecteur.listplaylist.AddItem txtnomfic.Text
    Else
        Lecteur.listplaylist.AddItem txtnomfic.Text & typefic
    End If
End If
        
'vérification de l'existence du fichier
For i = 0 To fichiers.ListCount - 1
    If Lecteur.listplaylist.List(Lecteur.listplaylist.ListCount - 1) = fichiers.List(i) Then
        ajoute_audio = True
        Exit For
    End If
Next i
        
End Function

Private Sub cmdannuler_Click()

Unload Explorateur 'décharge la feuille

End Sub

Private Sub cmdenregistrer_Click()
'enregistre une playlist

Dim numfic As String 'numéro de fichier
Dim i As Integer 'indice

'si le fichier existe déjà
For i = 0 To fichiers.ListCount - 1
    If txtnomfic.Text Like fichiers.List(i) Then
        If MsgBox("Ecrasez le fichier existant ?", vbYesNo + vbQuestion, "Sauvegarder") = vbYes Then
            Exit For
        Else
            Exit Sub
        End If
    End If
Next i

'création du fichier playlist
On Error GoTo Invalide 'nom du fichier incorrect
numfic = FreeFile

'ajout de l'extension si elle est absente et ouverture du fichier en écriture
If Right(txtnomfic.Text, 4) = ".jey" Then
    Open fichiers.Path & "\" & txtnomfic.Text For Output As #numfic
Else
    Open fichiers.Path & "\" & txtnomfic.Text & ".jey" For Output As #numfic
End If

'copie de la playlist dans le fichier
For i = 1 To Lecteur.listplaylist.ListCount - 1
    Print #numfic, Lecteur.listchemins.List(i - 1) & Lecteur.listplaylist.List(i)
Next i
Close numfic

Explorateur.Hide

Exit Sub

'nom du fichier incorrect
Invalide: MsgBox "Nom ou chemin du fichier incorrect !", vbCritical + vbOKOnly, "Erreur sauvegarde"
          On Error GoTo 0

End Sub

Private Sub cmdouvrir_Click()
'ouvre une playlist, un mp3, un wav ou un wma

Dim ligne As String 'ligne lue dans le fichier en entrée
Dim numfic As Integer 'numéro de fichier
Dim ajout As String
Dim carac As String
Dim i As Integer 'indice

'test du chemin
Select Case cbtype.ListIndex
    Case 0 '*.jey
        On Error GoTo Inexistant 'le fichier demandé n'existe pas
        numfic = FreeFile 'affectation d'un numéro de fichier
        
        'ajout de l'extension si elle est absente
        If Right(txtnomfic.Text, 4) = ".jey" Then
            Open fichiers.Path & "\" & txtnomfic.Text For Input As numfic
        Else
            Open fichiers.Path & "\" & txtnomfic.Text & ".jey" For Input As numfic
        End If
        
        'initialisation de la playlist
        Lecteur.listplaylist.Clear
        Lecteur.listplaylist.AddItem "---------------------------------------- Playlist ----------------------------------------"
        Lecteur.listplaylist.ListIndex = 0
    
        'chargement de la playlist sélectionnée
        While Not EOF(numfic)
            Line Input #numfic, ligne
            carac = ""
            i = 0
            While carac <> "\"
                i = i + 1
                carac = Right(Left(ligne, Len(ligne) - i), 1)
                ajout = Right(ligne, i)
            Wend
            Lecteur.listplaylist.AddItem ajout
            Lecteur.listchemins.AddItem Left(ligne, Len(ligne) - Len(ajout))
        Wend
        Close numfic
        On Error GoTo 0
        
    Case 1 '*.mp3
        'ajout d'un mp3 dans la playlist
        If Not ajoute_audio(".mp3") Then
            Lecteur.listplaylist.RemoveItem Lecteur.listplaylist.ListCount - 1
            GoTo Inexistant
        Else
            Lecteur.listchemins.AddItem fichiers.Path & "\"
        End If
        
    Case 2 '*.wav
        'ajout d'un wav dans la playlist
        If Not ajoute_audio(".wav") Then
            Lecteur.listplaylist.RemoveItem Lecteur.listplaylist.ListCount - 1
            GoTo Inexistant
        Else
            Lecteur.listchemins.AddItem fichiers.Path & "\"
        End If
        
    Case 3 '*.wma
        'ajout d'un wma dans la playlist
        If Not ajoute_audio(".wma") Then
            Lecteur.listplaylist.RemoveItem Lecteur.listplaylist.ListCount - 1
            GoTo Inexistant
        Else
            Lecteur.listchemins.AddItem fichiers.Path & "\"
        End If
        
End Select

Explorateur.Hide

Exit Sub

'le fichier demandé n'existe pas
Inexistant: MsgBox "Le fichier spécifié n'existe pas !", vbCritical + vbOKOnly, "Erreur d'ouverture"
            On Error GoTo 0

End Sub

Private Sub disc_Change()
'change le disque sélectionné

On Error GoTo disque 'le lecteur sélectionné n'est pas prêt
    dossiers.Path = disc.Drive
On Error GoTo 0

Exit Sub

'le lecteur sélectionné n'est pas prêt
disque:
MsgBox "Veuillez insérer un disque dans le lecteur !", vbOKOnly + vbCritical, "Erreur lecteur"

End Sub

Private Sub dossiers_Change()
'change le dossier sélectionné

fichiers.Path = dossiers.Path

End Sub

Private Sub dossiers_Click()
'change le dossier sélectionné

fichiers.Path = dossiers.List(dossiers.ListIndex)

End Sub

Private Sub dossiers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'affiche l'infobulle correspondant au dossier sélectionné

dossiers.ToolTipText = dossiers.List(dossiers.ListIndex)

End Sub

Private Sub fichiers_Click()
'change le fichier sélectionné

txtnomfic.Text = fichiers.List(fichiers.ListIndex)

End Sub

Private Sub fichiers_DblClick()

cmdouvrir_Click

End Sub

Private Sub fichiers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'affiche l'infobulle correspondant au fichier sélectionné

fichiers.ToolTipText = fichiers.List(fichiers.ListIndex)

End Sub

Private Sub Form_Load()

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name) 'changement du thème
dossiers.Path = disc.Drive

End Sub


