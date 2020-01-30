VERSION 5.00
Begin VB.Form Rechercher 
   Caption         =   "Rechercher par"
   ClientHeight    =   2160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdfermer 
      Cancel          =   -1  'True
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.OptionButton optprenom 
      Caption         =   "Pr�nom de &doubleur"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      ToolTipText     =   "Rechercher par pr�nom"
      Top             =   600
      Width           =   2055
   End
   Begin VB.OptionButton optprenom 
      Caption         =   "Pr�nom de &personnage"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      ToolTipText     =   "Rechercher par pr�nom"
      Top             =   960
      Width           =   2055
   End
   Begin VB.CommandButton cmdrechercher 
      Caption         =   "&Rechercher"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtrechercher 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Rechercher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force la d�claration des variables

Dim rst As New ADODB.Recordset

Private Sub cmdfermer_Click()

Unload Rechercher 'd�chargement de la form

End Sub

Private Sub cmdrechercher_Click()
'lancement des proc�dures stock�es de recherche par pr�nom

Dim rst2 As New ADODB.Recordset 'nouveau recordset
Dim rep As Long 'r�ponse � la msgbox

If optprenom(0).Value = True Then 'recherche pas pr�nom de seiyuu
    Dim cmd As New ADODB.Command 'd�claration d'un objet command
    Set cmd.ActiveConnection = Seiyuu.cn 'connection
    cmd.CommandText = "[rech_prenom_seiyuu]" 'proc�dure � executer
    
    'execution de la proc�dure stock�e
    'param�tre : prenom_seiyuu
    Set rst = cmd.Execute(, txtrechercher.Text)

'r�sultat obtenu
rechercheseiyuu:
    On Error GoTo mauvaisseiyuu
    rst2.Open "SELECT nom_seiyuu,prenom_seiyuu FROM SEIYUU WHERE num_seiyuu = " & rst!num_seiyuu, Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
    rep = MsgBox(rst2!nom_seiyuu & " " & rst2!prenom_seiyuu & " ?", vbYesNoCancel + vbQuestion, "R�sultat recherche")
    On Error GoTo 0
    
    If rep = vbYes Then 'confirmation par l'utilisateur
        Seiyuu.cbseiyuu.Text = rst2!nom_seiyuu & " " & rst2!prenom_seiyuu
        rst.Close
        rst2.Close
        Unload Rechercher
        Seiyuu.cbseiyuu_Click
    ElseIf rep = vbNo Then 'proposition d'un autre nom
        rst2.Close
        rst.MoveNext
        If Not rst.EOF Then
            GoTo rechercheseiyuu
        Else
            rst.Close
        End If
    ElseIf rep = vbCancel Then 'fermeture de la fen�tre
        rst.Close
        rst2.Close
    End If

Else 'recherche pas pr�nom de personnage

    'execution de la proc�dure stock�e
    'param�tre : prenom_perso
    rst.Open "rech_prenom_perso '" & txtrechercher.Text & "'", Seiyuu.cn, adOpenKeyset, adLockPessimistic

'r�sultat obtenu
rechercheperso:
    On Error GoTo mauvaisperso
    rst2.Open "SELECT nom_perso,prenom_perso FROM PERSO WHERE num_perso = " & rst!num_perso, Seiyuu.cn, adOpenForwardOnly, adLockPessimistic
    rep = MsgBox(rst2!nom_perso & " " & rst2!prenom_perso & " ?", vbYesNoCancel + vbQuestion, "R�sultat recherche")
    On Error GoTo 0
    
    If rep = vbYes Then 'confirmation par l'utilisateur
        Seiyuu.cbperso.Text = rst2!nom_perso & " " & rst2!prenom_perso
        rst.Close
        rst2.Close
        Unload Rechercher
        Seiyuu.cbperso_Click
    ElseIf rep = vbNo Then 'proposition d'un autre nom
        rst2.Close
        rst.MoveNext
        If Not rst.EOF Then
            GoTo rechercheperso
        Else
            rst.Close
        End If
    ElseIf rep = vbCancel Then 'fermeture de la fen�tre
        rst.Close
        rst2.Close
    End If
End If

Exit Sub

mauvaisseiyuu:
MsgBox "Aucun doubleur ne correspond au pr�nom saisi. Veuillez saisir un autre pr�nom.", vbOKOnly + vbExclamation, "Mauvais pr�nom"
rst.Close
On Error GoTo 0
Exit Sub

mauvaisperso:
MsgBox "Aucun personnage ne correspond au pr�nom saisi. Veuillez saisir un autre pr�nom.", vbOKOnly + vbExclamation, "Mauvais pr�nom"
rst.Close
On Error GoTo 0

End Sub

Private Sub Form_Load()

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name) 'changement du th�me
optprenom(0).Value = True 'pr�nom de seiyuu par d�faut
txtrechercher.Text = "pr�nom"
txtrechercher.SelLength = Len(txtrechercher.Text)

End Sub
