VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Lecteur 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Lecteur"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdfermer 
      Cancel          =   -1  'True
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Masquer le lecteur"
      Top             =   3480
      Width           =   495
   End
   Begin VB.ListBox listchemins 
      Height          =   450
      Left            =   3000
      TabIndex        =   17
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer tmrduree 
      Left            =   1320
      Top             =   2400
   End
   Begin VB.Timer tmrprogress 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2160
      Top             =   2400
   End
   Begin VB.CommandButton cmdouvrir 
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "Lecteur.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Ouvrir un fichier"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdstop 
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Arrêter"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdplaypause 
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdretour 
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Piste précédente"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CommandButton cmdavance 
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Piste suivante"
      Top             =   1440
      Width           =   495
   End
   Begin VB.CheckBox chkvolume 
      Height          =   375
      Left            =   2400
      Picture         =   "Lecteur.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.ListBox listplaylist 
      Height          =   1230
      ItemData        =   "Lecteur.frx":088C
      Left            =   120
      List            =   "Lecteur.frx":088E
      TabIndex        =   12
      Top             =   2040
      Width           =   4335
   End
   Begin VB.CommandButton cmdmoins 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Supprimer la piste sélectionnée"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdplus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ajouter le générique de la série sélectionnée à la playlist"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdnouveau 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      Picture         =   "Lecteur.frx":0890
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Effacer la playlist"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmdsauver 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      Picture         =   "Lecteur.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Sauvegarder la playlist"
      Top             =   3480
      Width           =   495
   End
   Begin MSComctlLib.ProgressBar progduree 
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.Slider sldvolume 
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      ToolTipText     =   "Volume"
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Max             =   100
      TickFrequency   =   10
   End
   Begin VB.Label lbltitre 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   840
      TabIndex        =   16
      ToolTipText     =   "Titre"
      Top             =   240
      UseMnemonic     =   0   'False
      Width           =   3615
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   495
      Left            =   3960
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label lblduree 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Lecteur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'force la déclaration des variables

Dim minutes As Integer
Dim secondes As Integer
Dim numpiste As Integer
Dim temps As Boolean
Dim sectotales As Double
Dim secrest As Double

Private Sub cmdavance_Click()
'passe à la piste suivante

If listplaylist.List(numpiste + 1) <> "" Then
    listplaylist.ListIndex = numpiste + 1
    listplaylist_DblClick
End If

End Sub

Private Sub cmdfermer_Click()
'masque le lecteur

Seiyuu.chkmedia.Value = 0

End Sub

Private Sub cmdouvrir_Click()
'charge la feuille "ouvrir"

Load Explorateur
With Explorateur
    .Caption = "Ouvrir"
    .cmdouvrir.Visible = True
    .cmdenregistrer.Visible = False
    .cmdouvrir.Default = True
    With .cbtype
        .Clear
        .AddItem "*.jey"
        .AddItem "*.mp3"
        .AddItem "*.wav"
        .AddItem "*.wma"
        .ListIndex = 0
    End With
    .Show vbModal
End With

End Sub

Private Sub cmdretour_Click()
'passe à la piste précédente

If listplaylist.List(numpiste - 1) <> "" Then
    listplaylist.ListIndex = numpiste - 1
    listplaylist_DblClick
End If

End Sub

Private Sub Form_Load()
'initialise les contrôles

cmdplaypause.ToolTipText = "Lire"
temps = True
chkvolume.ToolTipText = "Couper le son"
sldvolume.Value = 50 'initialisation du volume

'initialisation de la durée
lblduree.Caption = "00:00"
minutes = "00"
secondes = "00"

'initialisation de la playlist
listplaylist.AddItem "---------------------------------------- Playlist ----------------------------------------"
listplaylist.ListIndex = 0

Call Seiyuu.choix_theme(Seiyuu.cbtheme.ListIndex, Me.Name) 'changement du thème

End Sub

Private Sub chkvolume_Click()
'active/désactive le son

If chkvolume.Value = 1 Then
    WMP.settings.mute = True
    chkvolume.ToolTipText = "Remettre le son"
Else
    WMP.settings.mute = False
    chkvolume.ToolTipText = "Couper le son"
End If

End Sub

Private Sub cmdmoins_Click()
'retire une piste de la playlist

If listplaylist.ListIndex <> 0 Then
    listplaylist.ListIndex = listplaylist.ListIndex - 1
    listplaylist.RemoveItem listplaylist.ListIndex + 1
    listchemins.RemoveItem listplaylist.ListIndex
End If

End Sub

Private Sub cmdnouveau_Click()
'efface la playlist en cours

'nouvelle playlist
If listplaylist.ListCount > 1 Then
    If MsgBox("Effacer la playlist existante ?", vbYesNo + vbQuestion, "Nouvelle Playlist") = vbYes Then
        listplaylist.Clear
        listplaylist.AddItem "---------------------------------------- Playlist ----------------------------------------"
        listplaylist.ListIndex = 0
        listchemins.Clear
    End If
End If

End Sub

Private Sub cmdplaypause_Click()
'bouton play/pause

If WMP.playState = wmppsPlaying Then 'si le lecteur media est en cours de lecture
    WMP.Controls.pause
    tmrduree.Enabled = False 'interrompt le timer
    
    'boutons
    cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & cmdplaypause.Tag & "_play.jpg")
    cmdplaypause.ToolTipText = "Lire"
    
ElseIf WMP.playState = wmppsPaused Then 'si il est en pause
    WMP.Controls.play
    tmrduree.Enabled = True 'reprise du timer
    
    'boutons
    cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & cmdplaypause.Tag & "_pause.jpg")
    cmdplaypause.ToolTipText = "Mettre en pause"

's'il est arrêté et qu'aucune piste n'est sélectionnée
ElseIf listplaylist.ListIndex = -1 Or listplaylist.ListIndex = 0 Then
    If Seiyuu.cbanime.Text <> "" Then
        If Seiyuu.recup_titre(Seiyuu.cbanime.Text) <> "" Then
            WMP.URL = Command & "Génériques\" & Seiyuu.recup_titre(Seiyuu.cbanime.Text)
            lbltitre.Caption = WMP.currentMedia.Name
            
            'boutons
            cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & cmdplaypause.Tag & "_pause.jpg")
            cmdplaypause.ToolTipText = "Mettre en pause"
            
            'initialisation du timer
            tmrduree.Interval = 1000
            tmrduree.Enabled = True
        End If
    End If
    
'joue la piste sélectionnée
ElseIf listplaylist.ListIndex <> -1 And listplaylist.ListIndex <> 0 Then
    WMP.URL = Lecteur.listchemins.List(Lecteur.listplaylist.ListIndex - 1) & Lecteur.listplaylist.List(listplaylist.ListIndex)
    lbltitre.Caption = WMP.currentMedia.Name
    
    'initialisation du timer
    tmrduree.Interval = 1000
    tmrduree.Enabled = True
    
    'boutons
    cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & cmdplaypause.Tag & "_pause.jpg")
    cmdplaypause.ToolTipText = "Mettre en pause"
    
    numpiste = listplaylist.ListIndex

End If

End Sub

Private Sub cmdplus_Click()
'ajoute une piste à la playlist

If Seiyuu.cn.State = 1 And Seiyuu.cbanime.Text <> "" Then
    If Seiyuu.recup_titre(Seiyuu.cbanime.Text) <> "" Then
        listplaylist.AddItem Seiyuu.recup_titre(Seiyuu.cbanime.Text)
        listchemins.AddItem Command & "Génériques\"
    End If
End If

End Sub

Private Sub cmdsauver_Click()
'charge la feuille "sauvegarder"

Load Explorateur
If listplaylist.ListCount > 1 Then
    If MsgBox("Sauvegarder la playlist en cours ?", vbYesNo + vbQuestion, "Sauvegarder") = vbYes Then
        With Explorateur
            .Caption = "Sauvegarder"
            .cmdouvrir.Visible = False
            .cmdenregistrer.Visible = True
            .cmdenregistrer.Default = True
            .cbtype.Clear
            .cbtype.AddItem "*.jey"
            .cbtype.ListIndex = 0
            .txtnomfic.Text = "Playlist1.jey"
            .txtnomfic.SelLength = Len(.txtnomfic.Text)
            .Show vbModal
        End With
    End If
Else
    MsgBox "Aucune piste dans la playlist !", vbExclamation + vbOKOnly, "Sauvegarder"
End If

End Sub

Private Sub cmdstop_Click()
'arrête la lecture

WMP.Controls.Stop
cmdplaypause.Picture = LoadPicture(Command & "Boutons\" & cmdplaypause.Tag & "_play.jpg")
tmrprogress.Enabled = False
progduree.Value = 0
tmrduree.Enabled = False
lbltitre.Caption = ""
lblduree.Caption = "00:00"
WMP.URL = ""
minutes = "00"
secondes = "00"
sectotales = 0
secrest = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
'masque le lecteur

Seiyuu.chkmedia.Value = 0

End Sub

Private Sub lblduree_Click()
'inverse le défilement du temps

If temps = True Then
    temps = False
Else
    temps = True
End If

End Sub

Private Sub listplaylist_DblClick()
'joue la piste sélectionné

cmdstop_Click
cmdplaypause_Click

End Sub

Private Sub sldvolume_Scroll()
'modifie le son

WMP.settings.volume = sldvolume.Value

End Sub

Private Sub tmrprogress_Timer()
'récupère la position dans la piste pour faire défiler la ProgressBar

On Error Resume Next
progduree.Value = WMP.Controls.currentPosition
On Error GoTo 0

End Sub

Private Sub tmrduree_Timer()
'défilement du temps

If WMP.playState = wmppsPlaying Then
    'la propriété Max de la ProgressBar reçoit la durée de la piste
    progduree.Max = CInt(WMP.currentMedia.duration)
    
    'compte à rebours
    progduree.Max = progduree.Max - 1
End If

'défilement du temps
If secondes < 60 Then
    secondes = secondes + 1
    sectotales = sectotales + 1
Else
    secondes = "00"
    minutes = minutes + 1
End If

If temps = True Then
    'défilement du temps
    lblduree.Caption = Format(minutes, "00:") & Format(secondes, "00")
Else
    'compte à rebours
    secrest = progduree.Max - sectotales
    lblduree.Caption = Format(secrest / 60, "00:") & Format(secrest Mod 60, "00")
End If

tmrprogress.Enabled = True 'lance la progressbar

'passe à la piste suivante
If WMP.playState = wmppsStopped And listplaylist.List(listplaylist.ListIndex + 1) <> "" Then
    cmdavance_Click
ElseIf WMP.playState = wmppsStopped Then
    cmdstop_Click
End If

End Sub

