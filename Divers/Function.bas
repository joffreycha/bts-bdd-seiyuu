Attribute VB_Name = "funct1"
'Attribute VB_Name = "function"
Option Explicit
'jld20000416 fonctions de base  : outils vb

'jld20040628 f_lecver
Global nomg As String
Global nomd As String
Global nommasar As String
Global indexchemin As Integer ' Modif 04/10/95 pour spécification de chemin/masque

Global varmasque As String   'nom du masque par défaut (laprtf)
Global vbreseau As String   'jld20031207 client serveur : vbreseau

'FRED20030611
Global tabrwlap() As String
Global vartraaut As Integer 'jld20030512 repas traitement automatique

Global dossierwl As Integer 'utilisé dans creat_pa
Global tabarg() As String 'jld20020722

Global varmaxchp1 As Integer 'jld20020125
Global vardatvit As String 'jld20011210

Global gest_patientText As String 'jld20011125
Global varcry As Integer      'cryptage fichier 1:lapreso 2:secure 4:iid 8:chemin.lap 16:lapchem

Global varposreg As Integer   'position tableau règle
Global nbregle  As Integer
Global regle(300) As String

Global patient_selectionne As String 'jld20010218
Global varselpat As String 'jld20010302
Global trierliste As Integer  'jld20010311
Global nbselliste As Integer   'jld20040408 si 1 alors une seule sélection dans liste pla


Global intitule1 As String  'jld20010218
Global intitule2  As String 'jld20010218
Global intitule3  As String 'jld20010218
Global intitule4  As String 'jld20010218

Global varcom As Integer 'masque commun
Global varmen As Integer 'rafraichissement menu
Global varind As Integer 'rafraichissement menu
Global varact As String  'rafraichissement menu

Global varlistfont As Integer

   ' Define buttons.
Global Const MB_OK = 0
Global Const MB_OKCANCEL = 1
Global Const MB_YESNOCANCEL = 3
Global Const MB_YESNO = 4

   ' Define Icons.
Global Const MB_ICONSTOP = 16
Global Const MB_ICONQUESTION = 32

   ' Define other.
Global Const MB_ICONEXCLAMATION = 48
Global Const MB_ICONINFORMATION = 64
Global Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7

Global ShiftDown, AltDown, CtrlDown, vartxt
   
Global Const SHIFT_MASK = 1
Global Const CTRL_MASK = 2
Global Const ALT_MASK = 4

Global Const ATTR_NORMAL = 0        'Normal files
Global Const ATTR_HIDDEN = 2        'Hidden files
Global Const ATTR_SYSTEM = 4        'System files
Global Const ATTR_VOLUME = 8        'Volume label
Global Const ATTR_DIRECTORY = 16    'Directory

Global Const varspa = ";"  'séparateur de paramètres
Global Const varsep = ","  'séparateur de champ dans les fichiers lap
Global Const vardec = "."  'séparateur décimal pour les champs numériques lap
Global Const VarEuro = "6.55957"

Global varsepssv As String 'séparateur de sous valeurs par défaut = :

Global CrLf As String ' chr$(13) & chr$(10)

'ReDim Preserve tabint(Count + 10)  ' Resize the array.

Global varnumform As Integer ' 1 securite 2 lapsoc 4 lap_menu 8 masque

Global tabdes1() As String   'description champs masque

Global tabcry() As String    'jld20010913 tableau cryptage
Global tabchpsiz() As Integer
Global tabdocser() As String    'liste doc services
Global tabdocpra() As String    'liste doc praticiens
Global tabdocmod() As String    'liste doc modèles
Global tabdocide() As String    'liste doc utilisateurs
Global tabser() As String    'liste des services
Global tabmsq() As String    'propriétés du masque
Global tabdes() As String    'descriptions des champs du masque
Global tabchp() As String    'propriétés des champs du masque
Global tabint() As Long      'tableau entier temporaire
Global tablon() As Long      'tableau long temporaire
Global tabstr() As String    'tableau alpha temporaire
Global tabvar() As Variant   'tableau variant temporaire

Global tabtmp() As String
Global tabmem() As String
Global varmem1 As String
Global varmem2 As String
'jld20030721
'Global donnees(10, 500) As String
Global donnees(10, 400) As String
   'donnees(10,xxx) = mémorise les fenêtres ouvertes dans winlap
   'donnees(9,xxx)  = mémorise le contenu des champs du masque dans winlap
Global colstr As Long         'nombre de colonne du tableau tabstr

Global VarDecWin As String    'séparateur décimal windows

Global di$(30)                'contient les lignes de lapreso

'repertoires initiaux lapreso
Global varlaproot As String   '21
Global varlaphlp As String    '8
Global varlapdic As String    '18
Global varlaplap As String    '25
Global varlapmed As String    '1
Global varlapnum As String    '19
Global varlapvb As String     '20
Global varlappro As String     'jld20031207
Global varlapdi2 As String    '22
Global varlapspe As String    '26
Global varlapres As String    '12
Global varlaplet As String    '27 lettres
Global varlapleth As String    '7 lettres hors winlap
Global varlapsur As String    '30
Global varlaprap As String    '29 rapports rtf
Global varLaprep As String    '4  repas hopale
Global varLapreph As String   '5  repas hopale historique saisie lap
Global varLapdich As String   '6  dic repas hopale lap
Global varlaptrt As String    '24
Global varlaptmp As String

Global Varlapdos As String    'numero de dossier essai2
Global Varlapser As String    'service se

Global laparg As String       'arguments ligne de commande win_lap


'application
Global lapapp As String       'nom du programme
Global lapver As String       'version du programme
Global lappat As String       'numéro patient 1000
Global lapmmi As Integer      'dossier med millier alisutil
Global lapnmi As Integer      'dossier num millier alisutil
'repertoires courants

'jld20040621
Global varnbrver As Integer   'nbre de champ .ver
Global lapun As Integer       'premier démarrage de winlap
Global varmmi As Integer      'jld20040621 variable local millier masque principal
Global mapmmi As Integer      'jld20040621 variable local millier masque autres
Global lapsuf As String       'suffixe du masque principal : dossiers dans lapmed .std
Global varsuf As String       'suffixe du masque principal : dossiers dans lapmed
Global mapsuf As String       'suffixe des autres masques  : dossiers dans lapmed
Global mapmsq As String       'nom autre masque indexé
Global mapmed As String       'adresse repertoire dossier autre masque indexé
Global tabverm() As String    'fichier .ver masque principal
Global tabvera() As String    'fichier .ver autres masques

Global lapmsq As String       'dossier des fichiers temporaires masque : laptmp\winlap\1\ à laptmp\winlap\9\ le 0 est pour winlap
Global lapessai2 As String    'mémorise le dossier patient de winlap si masqnorm
Global laproot As String      '21
Global laphlp As String       '8
Global lapdic As String       '18
Global laplap As String       '25
Global lapmed As String       '1
Global lapnum As String       '19
Global lapvb As String        '20
Global lappro As String        'jld20031207
Global lapdi2 As String       '22
Global lapspe As String       '26
Global lapres As String       '12
Global laplet As String       '27 lettres
Global lapleth As String       '7 lettres hors winlap
Global lapsur As String       '30
Global laprap As String       '29 rapports rtf
Global laprep As String       '4  repas hopale
Global lapreph As String      '5  repas hopale historique saisie lap
Global lapdich As String      '6  repas hopale historique saisie lap
Global laptrt As String       '24
Global laptmp As String
'jld20040909
Global laprem As String       'repertoire temporaire qui sera effacé à la fin du programme
Global laptra As String       'répertoire de trace des problèmes
Global lapchp As Integer
Global laptxt As String
Global lapchp2 As Integer
Global laptxt2 As String
'fred20030429
Global lapbac As String
Global lapana As String

Global lapage As String       'info calculé   : age patient
Global lapnom As String       'info num   : nom usuel
Global lappre As String       'info num   : prénom
Global lapnai As String       'info num   : date de naissance
Global lapsex As String       'info num   : sexe
Global lapdos As String       'info num   : numéro de dossier essai2
Global lapide As String       'info login : identité utilisateur
Global lapmet As String       'info login : metier utilisateur   jld20030512
Global lapniv As String       'info login : niveau utilisateur   jld20030512
Global Const lapdosrep = "20" 'info       : numéro de dossier repas   : 20 hopale
Global Const lapdoshop = "998" 'info      : numéro de dossier hopital : 998


Global lapnna As String       'info num : nom de naissance
Global lapdoc As String       'info num : dernier numéro de doc

Global lapser As String       'service se
Global Lapser2 As String      'choix service se dans lap_menu
Global Laphos As String       'hospit

'jld20031207 Global lappro As String       'dernier programme utilisé
Global lapdat As String       'date en cours
Global lapheu As String       'heure en cours

Global laplog As String       'login code service
Global lappas As String       'passe code perso
Global lapper As Integer      'niveau permission
Global lapeta As String       'etat de travail : nouveau modif etc

'Fred20011023 Pour multijour
Global fgdate() As String     'Tableau des dates des configs a prendre en compte
Global fgtypedate As String   'Prise en charge multijou fgtypedate = MULTIJOUR si oui
Global fgnbjour As Integer    'Nombre de jour dont la config est a prendre en compte
Global fgjour As String       'Nom du jour actuel de la semaine
''Parametrage du nombre mlax de config a sauvegardees
Global nbjlu As Integer    'Lundi
Global nbjma As Integer    'Mardi
Global nbjme As Integer    'Mercredi
Global nbjje As Integer    'Jeudi
Global nbjve As Integer    'Vendredi
Global nbjsa As Integer    'Samedi
Global nbjdi As Integer    'Dimanche
Global nbjtot As Integer   'Nombre max de config a conserver ds l'archivage

'FRED20011026 fonction mk_dir
Global pb As Integer    'pb = 1 si la fct ne s'est pas correctement terminee, sinon pb = 0

'FRED20030324
Global globnote As String

Function ansi_ascii(varstr1)
    'ansi_ascii = varstr1
    'Exit Function
    'On remplace les caractères ansi spéciaux par leur équivalent ascii
    Dim carin As String
    Dim varstr As String

    Static c1, c2, c3, c4, accent, ansi As String
    Static car As String
    Static i, ins As Integer

    varstr = varstr1
    
    c1 = Chr$(132) + Chr$(148) + Chr$(129) + Chr$(225) + Chr$(142) + Chr$(153)
    c2 = Chr$(154) + Chr$(128) + Chr$(130) + Chr$(131) + Chr$(133) + Chr$(134) + Chr$(135)
    c3 = Chr$(136) + Chr$(137) + Chr$(138) + Chr$(139) + Chr$(140) + Chr$(143) + Chr$(144)
    c4 = Chr$(147) + Chr$(150) + Chr$(151)
    accent = c1 + c2 + c3 + c4
    ansi = "äöüßÄÖÜÇéâàåçêëèïîÅÉôûù"
    carin = ""
    For i = 1 To Len(varstr)
        car = Mid$(varstr, i, 1)
        If Asc(car) > 127 Then
            ins = InStr(ansi, car)
            If ins = 0 Then
                carin = car
            End If
        End If
    Next i
    Err = 0 'jld20040816 err 5 si a$ vide : asc
    
    For i = 1 To Len(ansi)
        Do
            ins = InStr(varstr, Mid$(ansi, i, 1))
            If ins = 0 Then
                Exit Do
            Else
                Mid$(varstr, ins, 1) = Mid$(accent, i, 1)
            End If
        Loop
    Next i
    
    ansi_ascii = varstr
    
End Function

Function ascii_ansi(texte) As String
    'On remplace les caractères ascii spéciaux par leur équivalent ansi
    Static c1, c2, c3, c4, accent, ansi As String
    Static i, ins As Integer

    c1 = Chr$(132) + Chr$(148) + Chr$(129) + Chr$(225) + Chr$(142) + Chr$(153)
    c2 = Chr$(154) + Chr$(128) + Chr$(130) + Chr$(131) + Chr$(133) + Chr$(134) + Chr$(135)
    c3 = Chr$(136) + Chr$(137) + Chr$(138) + Chr$(139) + Chr$(140) + Chr$(143) + Chr$(144)
    c4 = Chr$(147) + Chr$(150) + Chr$(151) + Chr$(253) + Chr$(241)
    accent = c1 + c2 + c3 + c4
    ansi = "äöüßÄÖÜÇéâàåçêëèïîÅÉôûù²±"
    For i = 1 To Len(accent)
        Do
            ins = InStr(texte, Mid$(accent, i, 1))
            If ins = 0 Then
                Exit Do
            Else
                Mid$(texte, ins, i) = Mid$(ansi, i, 1)
            End If
        Loop
    Next i
    ascii_ansi = texte
End Function

Function f_1000(varstr1 As String, varstr2 As String) As String
'traitement des adresses avec modulo 1000 ( accès rapide au répertoire )
'varstr1 = N° Dossier ex : 45029 = 45\45029  attention varstr1 peut être = 1000.std pour type 1
'varstr1 = N° Dossier:masque ex : 45029:ADM donne masque\45029  attention varstr1 peut être = 1000.std pour type 2
'varstr2 = Adresse répertoire avec AntéSlah ex : lapmed ou c:\med\ pour type 1 : sta
'retour = adresse dossier : millier\numéro ou masque\numéro
'si code lapmmi = 2 alors AIDER : med = med\masque\dossier.std

   Dim varvar As Variant
   Dim vardos As String
   Dim varmsq As String
   Dim vardir As String
   Dim varadr As String
   Dim varstr As String
   Dim varrep As String
   Dim varlon As Long
   Dim varint As Integer
   Dim varfin As String
   
   f_1000 = varstr1

   vardos = Trim$(varstr1)
   vardir = Trim$(varstr2)
   varmsq = ""
   
   If InStr(varstr1, ":") > 0 Then
      vardos = Trim$(f_champ(varstr1, ":", 1))
      varmsq = Trim$(f_champ(varstr1, ":", 2))
   End If

   'jld20021112 aider
   If varmsq <> "" Then
      f_1000 = varmsq & "\" & vardos
      Exit Function
   End If
   
   varint = InStr(vardos, ".")
   If varint > 0 Then
      vardos = Trim$(Left$(vardos, varint - 1))
   End If
   
   varvar = vardos
   If Not IsNumeric(varvar) Then
      f_1000 = varstr1
      Exit Function
   End If

   varlon = Val(vardos)

   If varlon = 0 Then
      f_1000 = varstr1
      Exit Function
   End If
   
   varadr = Trim$(Str$(Int(varlon / 1000)))
   
   If Trim$(vardir) <> "" Then
      If Right$(Trim$(vardir), 1) <> "\" Then
          vardir = vardir + "\"
      End If
      varrep = f_mkdir(vardir + varadr, "", 0)
   End If

   f_1000 = varadr + "\" + Trim$(varstr1)


End Function

Function f_adr(varstr1 As String) As String
'traitement des adresses avec variables winlap
'varstr1 = adresse

   Dim varstr As String
   Dim varrep As String
   Dim varadr As String

   varstr = Trim$(UCase$(varstr1))

   'jld20021217 élimination des "
   varstr = f_remplace(varstr, """", "", "T", 1)

   If Left$(varstr, 2) = "\\" Then
      varstr = f_remplace(varstr, "\\", "\", "T", 1)
      varstr = "\" & varstr
   Else
      varstr = f_remplace(varstr, "\\", "\", "T", 1)
   End If

   varrep = f_champ(varstr, "\", 1)
   varadr = Mid$(varstr, Len(varrep) + 2)

'      varlaphlp = laphlp
'      varlapdic = lapdic
'      varlaplap = laplap
'      varlapmed = lapmed
'      varlapnum = lapnum
'      varlapvb = lapvb
'      varlaproot = laproot
 '     varlapdi2 = lapdi2
 '     varlapres = lapres
 '     varlapspe = lapspe
 '     varlaplet = laplet
 '     varlapleth = lapleth
 ''     varlapsur = lapsur
 '     varlaprap = laprap
 '     varLaprep = laprep
 '     varLapreph = lapreph
 '     varLapdich = lapdich
 '     varlaptrt = laptrt
  '    varlaptmp = laptmp
  
   Select Case Trim$(UCase$(varrep))
   Case "LAPROOT"
   'jld20040413 masqnorm multienvironnement yyy
      f_adr = laproot & varadr
   'jld20040909
   Case "LAPREM"
      f_adr = laprem & varadr
   Case "LAPTMP"
      f_adr = laptmp & varadr
   'jld20031020
   Case "LAPHLP"
      f_adr = laphlp & varadr
   Case "LAPRES"
      f_adr = lapres & varadr
   Case "LAPTRA"
      f_adr = laptra & varadr
   Case "LAPDIC"
      f_adr = lapdic & varadr
   Case "LAPDICH"
      f_adr = lapdich & varadr
   Case "LAPLAP"
      f_adr = laplap & varadr
   Case "LAPMED"
      f_adr = lapmed & varadr
   Case "LAPNUM"
      f_adr = lapnum & varadr
   Case "LAPDI2"
      f_adr = lapdi2 & varadr
   Case "LAPRAP"
      f_adr = laprap & varadr
   Case "LAPREP"
      f_adr = laprep & varadr
   Case "LAPREPH"
      f_adr = lapreph & varadr
   Case "LAPLET"
      f_adr = laplet & varadr
   Case "LAPLETH"
      f_adr = lapleth & varadr
   Case "LAPSPE"
      f_adr = lapspe & varadr
   Case "LAPVB"
      f_adr = lapvb & varadr
   'jld20031207
   Case "LAPPRO"
      f_adr = lappro & varadr
   Case Else
      f_adr = varstr1
   End Select
   
   'jld20041020
   varstr = ""
   varstr = f_adr
   
   If InStr(1, varstr, " LAPROOT\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPROOT\", " " & laproot, "T", 1)
   End If
   If InStr(1, varstr, " LAPREM\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPREM\", " " & laprem, "T", 1)
   End If
   If InStr(1, varstr, " LAPTMP\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPTMP\", " " & laptmp, "T", 1)
   End If
   If InStr(1, varstr, " LAPHLP\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPHLP\", " " & laphlp, "T", 1)
   End If
   If InStr(1, varstr, " LAPRES\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPRES\", " " & lapres, "T", 1)
   End If
   If InStr(1, varstr, " LAPTRA\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPTRA\", " " & laptra, "T", 1)
   End If
   If InStr(1, varstr, " LAPDIC\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPDIC\", " " & lapdic, "T", 1)
   End If
   If InStr(1, varstr, " LAPDICH\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPDICH\", " " & lapdich, "T", 1)
   End If
   If InStr(1, varstr, " LAPLAP\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPLAP\", " " & laplap, "T", 1)
   End If
   If InStr(1, varstr, " LAPMED\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPMED\", " " & lapmed, "T", 1)
   End If
   If InStr(1, varstr, " LAPNUM\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPNUM\", " " & lapnum, "T", 1)
   End If
   If InStr(1, varstr, " LAPDI2\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPDI2\", " " & lapdi2, "T", 1)
   End If
   If InStr(1, varstr, " LAPRAP\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPRAP\", " " & laprap, "T", 1)
   End If
   If InStr(1, varstr, " LAPREP\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPREP\", " " & laprep, "T", 1)
   End If
   If InStr(1, varstr, " LAPREPH\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPREPH\", " " & lapreph, "T", 1)
   End If
   If InStr(1, varstr, " LAPLET\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPLET\", " " & laplet, "T", 1)
   End If
   If InStr(1, varstr, " LAPLETH\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPLETH\", " " & lapleth, "T", 1)
   End If
   If InStr(1, varstr, " LAPSPE\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPSPE\", " " & lapspe, "T", 1)
   End If
   If InStr(1, varstr, " LAPVB\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPVB\", " " & lapvb, "T", 1)
   End If
   If InStr(1, varstr, " LAPPRO\", 1) > 0 Then
      varstr = f_remplace(varstr, " LAPPRO\", " " & lappro, "T", 1)
   End If

   f_adr = varstr
   
End Function

Function f_aider(varstr1 As String)
'traitement des adresses Aider avec masque : type 2 lapmmi
'concatenation des dossier en un seul dans med
'si le fichier laprtfmq.lap existe dans vb : seul les masques du fichier sont traités
'varstr1 = N° Dossier   attention varstr1 peut être = 1000.std
'retour = OK
'si code lapmmi = 2 alors AIDER : med = med\masque\dossier.std

   Dim varerr As Integer
   Dim varvar As Variant
   Dim varfic As String
   Dim vardos As String
   Dim varmsq As String
   Dim vardir As String
   Dim varadr As String
   Dim varstr As String
   Dim varrep As String
   Dim varlon As Long
   Dim varint As Integer
   Dim varfin As String
   Dim varnbr As Integer
   Dim nf As Integer
   Dim nf2 As Integer
   Dim i As Integer
   Dim a$
   Dim find As Integer
   Dim nbrmas As Integer
   Dim tabmag() As String
   ReDim tabmag(0)
   Dim tabmas() As String
   ReDim tabmas(0)
   
   
   f_aider = ""

   vardos = Trim$(varstr1)
   vardir = ""
   varmsq = ""
   
   If InStr(vardos, ".") = 0 Then
      vardos = vardos & ".std"
   End If
   
   'varvar = vardos
   'If Not IsNumeric(varvar) Then
   '   f_1000 = varstr1
   '   Exit Function
   'End If

   On Error Resume Next
   varstr = ""
   varstr = Dir$(lapmed & "HOIND01", 16)
   If varstr = "" Then
      varstr = f_mkdir(lapmed & "HOIND01", "", 0)
   End If
   
   On Error Resume Next
   varstr = ""
   varstr = Dir$(lapmed & "HOIND02", 16)
   If varstr = "" Then
      varstr = f_mkdir(lapmed & "HOIND02", "", 0)
   End If
   
   On Error Resume Next
   varstr = ""
   'varstr = Dir$(lapmed & "HOIND03", 16)
   varstr = Dir$(lapmed & varmasque, 16)
   If varstr = "" Then
      'varstr = f_mkdir(lapmed & "HOIND03", "", 0)
      varstr = f_mkdir(lapmed & varmasque, "", 0)
   End If

   On Error Resume Next
   varstr = f_mkdir(lapmed & "TMPMSQ", "", 0)
   On Error Resume Next
   Kill lapmed & "TMPMSQ" & "\" & vardos

   'jld20021120 test existance fichier liste masque à consulter
   varerr = 0
   find = False
   nbrmas = 0
   On Error Resume Next
   nf = FreeFile
   Close #nf: Open lapvb & "laprtfmq.lap" For Input As #nf
   If Err = 0 Then
      Do While Not EOF(nf)
         Line Input #nf, a$
         a$ = Trim$(UCase$(a$))
         If Left$(a$, 1) = "[" And find = True Then
            Exit Do
         End If
         If a$ <> "" Then
            Select Case Trim$(UCase$(lapapp))
            Case "LAPRTF"
               If find = True Or a$ = "[LAPRTF]" Then
                  find = True
                  If a$ <> "[LAPRTF]" Then
                     nbrmas = nbrmas + 1
                     ReDim Preserve tabmas(nbrmas)
                     tabmas(nbrmas) = a$
                  End If
               End If
               
            Case "LAPDOC"
               If find = True Or a$ = "[LAPDOC]" Then
                  find = True
                  If a$ <> "[LAPDOC]" Then
                     nbrmas = nbrmas + 1
                     ReDim Preserve tabmas(nbrmas)
                     tabmas(nbrmas) = a$
                  End If
               End If
               
            Case "AGENDACB"
               If find = True Or a$ = "[AGENDACB]" Then
                  find = True
                  If a$ <> "[AGENDACB]" Then
                     nbrmas = nbrmas + 1
                     ReDim Preserve tabmas(nbrmas)
                     tabmas(nbrmas) = a$
                  End If
               End If
               
            Case "WIN_LAP"
               If find = True Or a$ = "[WIN_LAP]" Then
                  find = True
                  If a$ <> "[WIN_LAP]" Then
                     nbrmas = nbrmas + 1
                     ReDim Preserve tabmas(nbrmas)
                     tabmas(nbrmas) = a$
                  End If
               End If

            Case Else
            End Select
         End If
      Loop
   End If
   'jld20030106
   Close #nf

   If find = True And UBound(tabmas) > 0 Then

      On Error Resume Next
      varint = 0
      varnbr = 0
      varadr = ""
      vardir = ""
      For i = 1 To UBound(tabmas)
         varerr = 0
         vardir = tabmas(i)
         If Trim$(UCase$(vardir)) <> "TMPMSQ" Then
            varfic = ""
            varint = 0
            On Error Resume Next
            'varint = FileLen(lapmed & vardir & "\" & vardos)
            nf2 = FreeFile
            Close #nf2: Open lapmed & vardir & "\" & vardos For Input Lock Write As #nf2
            If Err Then
               varerr = Err
            Else
               varerr = 0
            End If
            'jld20031224 Unlock #nf2
            Close #nf2

            'If varint <> 0 And Err = 0 Then
            If varerr = 0 Then
               varnbr = varnbr + 1
               If varnbr = 1 Then
                  On Error Resume Next
                  FileCopy lapmed & vardir & "\" & vardos, lapmed & "tmpmsq" & "\" & vardos
               Else
                  varstr = fg_cpfic(lapmed & vardir & "\" & vardos, lapmed & "tmpmsq" & "\" & vardos, "APP::N")
               End If
            End If
         End If
      Next i
   
   Else

      On Error Resume Next
      varint = 0
      varnbr = 0
      varadr = ""
      vardir = ""
      vardir = Dir$(lapmed & "*.*", 16)
      Do While vardir <> ""
         If vardir <> "." And vardir <> ".." And Trim$(UCase$(vardir)) <> "TMPMSQ" Then
            varfic = ""
            varint = 0
            On Error Resume Next

            'varint = FileLen(lapmed & vardir & "\" & vardos)
            nf2 = FreeFile
            Close #nf2: Open lapmed & vardir & "\" & vardos For Input Lock Write As #nf2
            If Err Then
               varerr = Err
            Else
               varerr = 0
            End If
            'jld20031224 Unlock #nf2
            Close #nf2

            'If varint <> 0 And Err = 0 Then
            If varerr = 0 Then

               varnbr = varnbr + 1
               If varnbr = 1 Then
                  On Error Resume Next
                  FileCopy lapmed & vardir & "\" & vardos, lapmed & "tmpmsq" & "\" & vardos
               Else
                  varstr = fg_cpfic(lapmed & vardir & "\" & vardos, lapmed & "tmpmsq" & "\" & vardos, "APP::N")
               End If
            End If
         End If
         On Error Resume Next
         vardir = Dir$
      Loop
   End If

   'err
   If UBound(tabmas) > 0 Then
      f_aider = "OK-"
   Else
      f_aider = "OK"
   End If

End Function

Function f_alisvar(varstr1 As String) As String
'lecture fichier alislap.ini PATH WLREPAS à partir de alisutil.txt
'varstr1 = adresse alisutil.txt par défaut même repertoire que appli
   
   Dim vardir As String
   Dim varfic As String
   Dim varficres As String
   Dim varstr As String
   Dim varrep As String
   Dim varwindir As String
   Dim varalis As String
   Dim varutil As String
   Dim varser As String
   ReDim tabser(0)
   'jld20011107 gestion doc
   ReDim tabdocser(0)
   ReDim tabdocpra(0)
   ReDim tabdocmod(0)
   ReDim tabdocide(0)
   Dim varnbr As Integer
   Dim varnbrser As Integer

   Dim varint As Integer
   Dim find As Integer
   Dim nf As Integer
   Dim i As Integer
   Dim j As Integer
   Dim k As Integer
   Dim a$
   Dim b$

   f_alisvar = ""

   varstr = f_const()

   'curdir
   varwindir = ""
   varstr = ""
   varutil = Trim$(varstr1)
   If varutil = "" Then varutil = "alisutil.txt"
    
'lecture alisutil.txt : adresse alislap.ini

   varalis = ""
   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varutil For Input As #nf
   If Err = 0 Then
      Do While Not EOF(nf)
         'jld20040816
         On Error Resume Next
         Line Input #nf, a$
            If Err Then 'jld20011203
               MsgBox Str(Err) & " ERREUR WHILE alisvar 1"
               Exit Do
            End If
         
         If Trim$(UCase$(a$)) = "[ALISLAP]" Then
            varint = 0
            varalis = ""
            Do While Not EOF(nf)
               'jld20040816
               On Error Resume Next
            
               Line Input #nf, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE alisvar 2"
                  Exit Do
               End If

               If Trim$(a$) = "" Or Left$(Trim$(a$), 1) = "[" Then
                  Exit Do
               End If
               
               varstr = Trim$(UCase$(f_champ(a$, "=", 1)))
               If varstr = "PATH" Then
                  varalis = Trim$(UCase$(f_champ(a$, "=", 2)))
                  'jld20020113
                  vardir = Trim$(CurDir$)
                  If Right$(vardir, 1) = "\" Then
                     vardir = Left$(vardir, Len(vardir) - 1)
                  End If
                  varint = f_remplit_str(varalis, "\")
                  If tabstr(1) = ".." And Len(varalis) >= 3 Then
                     vardir = Trim$(f_dirname(vardir))
                      varalis = vardir & Right$(varalis, Len(varalis) - 3)
                  End If
                  If tabstr(1) = "." And Len(varalis) >= 3 Then
                     vardir = Trim$(vardir) & "\"
                      varalis = vardir & Right$(varalis, Len(varalis) - 3)
                  End If
               End If
               
               'jld20031207 vbreseau
               If varstr = "VBRESEAU" Then
                  vbreseau = Trim$(f_champ(a$, "=", 2))
                    'jld20040105
                    If Trim$(UCase$(Left$(vbreseau, 4))) = "PWD\" Then
                      vbreseau = CurDir$ & Mid$(vbreseau, 4)
                    End If
               End If

               'jld20040107 multibase
               If varstr = "MULTIBASE" Then
                  varstr = Trim$(f_champ(a$, "=", 2))
                  varstr = Trim$(UCase$(Left$(varstr, 1)))
                  If varstr = "O" Or varstr = "1" Then
                      On Error Resume Next
                      varstr = ""
                      varstr = Dir$("c:\vb\lapvar.txt", 0)
                      If Trim$(UCase$(varstr)) = "LAPVAR.TXT" And Err = 0 Then
                           On Error Resume Next
                           varstr = ""
                           varstr = Dir$("c:\vb\lapreso", 0)
                           If Trim$(UCase$(varstr)) = "LAPRESO" And Err = 0 Then
                              varficres = "c:\vb\lapreso"
                              GoTo suitealisvar
                           End If
                      End If
                   End If
               End If

            Loop
         End If
         'jld20010925 cryptage    SAUT_LIGNE_PAT=1
         If Trim$(UCase$(a$)) = "[LAPRESO]" Then
            varint = 0
            varcry = 0
            Do While Not EOF(nf)
               'jld20040816
               On Error Resume Next
            
               Line Input #nf, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE alisvar 3"
                  Exit Do
               End If
               
               If Trim$(a$) = "" Or Left$(Trim$(a$), 1) = "[" Then
                  Exit Do
               End If
               varstr = Trim$(UCase$(f_champ(a$, "=", 1)))
               If varstr = "CRY" Then
                  varcry = Val(Trim$(UCase$(f_champ(a$, "=", 2))))
               End If
            Loop
         End If

      Loop
   Else
      If Trim$(varstr1) <> "" Then
         MsgBox "ERREUR LECTURE fichier : " & f_basname(varstr1)
      End If
      Err = 0
   End If
   Close #nf
   On Error GoTo 0

   If Trim$(varalis) <> "" Then
      varfic = varalis
   Else
      
      varwindir = Environ("windir")
   
      If varwindir = "" Then
         'jld20021024
         On Error Resume Next
         varstr = "" 'jld
         varstr = Dir("c:\windows", 16)
         If varstr = "" Then
            'jld20021024
            On Error Resume Next
            varstr = Dir("c:\winnt", 16)
         End If
         If varstr = "" Then
            MsgBox "impossible de trouver le répertoire de Windows!"
            End
         Else
            varwindir = varstr
         End If
   
      End If
   
      varfic = varwindir & "\alislap.ini"
   End If

   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic For Input As #nf
   If Err Then
      On Error Resume Next
      Open varfic For Output As #nf
      If Err Then
         MsgBox Str(Err) & " erreur ecriture : " & f_basname(varfic)
         Close #nf
         Err = 0
         End
      Else
         Print #nf, "[WIN_LAP]"
         Print #nf, "PATH=C:\VB\LAPRESO"
         Print #nf, "[WLREPAS]"
         Print #nf, "PATH=C:\VB\LAPRESO"
         Close #nf
         varficres = "C:\VB\LAPRESO"
      End If
   Else
      find = False
      varrep = ""
      varficres = ""
      Do While Not EOF(nf)
         'jld20040816
         On Error Resume Next
      
         Line Input #nf, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE alisvar 4"
            Err = 0
            Exit Do
         End If

         'jld20011107 gestion doc
         If Trim$(UCase$(a$)) = "[WIN_LAP]" Then
               varnbr = 0
               Do While Not EOF(nf)
                  'jld20040816
                  On Error Resume Next
               
                  Line Input #nf, b$
                  If Err Then 'jld20011203
                     MsgBox Str(Err) & " ERREUR WHILE alisvar 6"
                     Err = 0
                     Exit Do
                  End If

                     If Left$(Trim(b$), 1) <> "[" Then

                           'jld géré dans [WLREPAS]
                           'If InStr(1, b$, "PATH=", 1) <> 0 And ficreso = "" Then
                           '   varrep = Trim$(f_champ(b$, "=", 2))
                           '   ficreso = Trim$(f_champ(varrep, varsep, 1))
                           'End If
                           
                           If Trim$(UCase$(f_champ(b$, "=", 1))) = "SERVICES" Then
                              k = 0
                              varrep = Trim$(f_champ(b$, "=", 2))
                              varnbrser = 0
                              varnbrser = f_remplit_str(varrep, varsep)
                              If varnbrser > 0 Then
                                 For i = 1 To varnbrser
                                    'jld20020715 vérification doublon
                                    find = False
                                    For j = 1 To UBound(tabser)
                                       If Trim$(UCase$(f_champ(tabser(j), " ", 1))) = Trim$(UCase$(f_champ(tabstr(i), " ", 1))) Then
                                          find = True
                                          Exit For
                                       End If
                                    Next j
                                    If find = False Then
                                       k = k + 1
                                       ReDim Preserve tabser(k)
                                       tabser(k) = Trim$(UCase$(tabstr(i)))
                                    End If
                                 Next i
                              End If
                           End If

                        'jld20011107 gestion doc
                           If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCSER" Then
                              varrep = Trim$(f_champ(b$, "=", 2))
                              varnbr = 0
                              varnbr = f_remplit_str(varrep, varsep)
                              If varnbr > 0 Then
                                 For i = 1 To varnbr
                                    ReDim Preserve tabdocser(i)
                                    tabdocser(i) = Trim$(UCase$(tabstr(i)))
                                 Next i
                              End If
                           End If

                           If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCPRA" Then
                              varrep = Trim$(f_champ(b$, "=", 2))
                              varnbr = 0
                              varnbr = f_remplit_str(varrep, varsep)
                              If varnbr > 0 Then
                                 For i = 1 To varnbr
                                    ReDim Preserve tabdocpra(i)
                                    tabdocpra(i) = Trim$(UCase$(tabstr(i)))
                                 Next i
                              End If
                           End If

                           If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCMOD" Then
                              varrep = Trim$(f_champ(b$, "=", 2))
                              varnbr = 0
                              varnbr = f_remplit_str(varrep, varsep)
                              If varnbr > 0 Then
                                 For i = 1 To varnbr
                                    ReDim Preserve tabdocmod(i)
                                    tabdocmod(i) = Trim$(UCase$(tabstr(i)))
                                 Next i
                              End If
                           End If

                           If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCIDE" Then
                              varrep = Trim$(f_champ(b$, "=", 2))
                              varnbr = 0
                              varnbr = f_remplit_str(varrep, varsep)
                              If varnbr > 0 Then
                                 For i = 1 To varnbr
                                    ReDim Preserve tabdocide(i)
                                    tabdocide(i) = Trim$(UCase$(tabstr(i)))
                                 Next i
                              End If
                           End If


                     Else
                           a$ = b$
                           Exit Do
                     End If
               Loop
         End If
      Loop
   End If
   Close #nf
   On Error GoTo 0

   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic For Input As #nf
   If Err Then
         MsgBox Str(Err) & " ERREUR LECTURE fichier : " & f_basname(varfic)
         Close #nf
         Err = 0
         End
   Else
      find = False
      varrep = ""
      varficres = ""
      Do While Not EOF(nf)
         'jld20040816
         On Error Resume Next
      
         Line Input #nf, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE alisvar 7"
            Err = 0
            Exit Do
         End If
         
         If Trim$(UCase$(a$)) = "[WLREPAS]" Then
            find = True
            Do While Not EOF(nf)
               'jld20040816
               On Error Resume Next
            
               Line Input #nf, b$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE alisvar 8"
                  Err = 0
                  Exit Do
               End If

               If Left$(Trim(b$), 1) <> "[" Then
                  'on garde la premiere ligne de alislap.ini
                  If InStr(1, b$, "PATH=", 1) <> 0 And varficres = "" Then
                     varficres = Trim$(f_champ(b$, "=", 2))
                     varficres = Trim$(f_champ(varficres, varsep, 1))
                  End If
                  
                  'jld20011107 les services sont définis dans [WIN_LAP]
                  'If Trim$(UCase$(f_champ(b$, "=", 1))) = "SERVICES" Then
                  '   varstr = Trim$(f_champ(b$, "=", 2))
                  '   varint = f_remplit_str(varstr, varsep)
                  '   For i = 1 To varint
                  '      ReDim Preserve tabser(i)
                  '      tabser(i) = tabstr(i)
                  '   Next i
                  'End If
               Else
                  'jld20011107
                  a$ = b$
                  Exit Do
               End If
            Loop
         End If
      Loop
      Close #nf
      On Error GoTo 0

      If find = False Then
         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Append As #nf
         If Err Then
            MsgBox Str(Err) & " erreur ecriture : " & f_basname(varfic)
            Close #nf
            Err = 0
            End
         Else
        
            Print #nf, "[WLREPAS]"
            Print #nf, "PATH=C:\VB\LAPRESO,WLREPAS"
            Close #nf
            varficres = "C:\VB\LAPRESO"
         End If
      End If
   End If
   Close #nf
   On Error GoTo 0
   
suitealisvar:
    
   varstr = ""
   varstr = f_lapvar(varficres)

   If varstr <> "OK" Then
      MsgBox "ERREUR LECTURE fichier : " & f_basname(varficres)
      End
   End If

   f_alisvar = "OK"

finlapreso:

   

End Function

Function f_annee(varstr1 As String, varlen As Integer) As String
'varstr1 = annee yy ou yyyy
'varlen  = nombre de caractère année
'resultat = f_annee année convertie sur le nombre de caractère varlen
   Dim varstr As String

   f_annee = varstr1

   varstr = Right$(varstr1, varlen)
   
   If Len(varstr1) = 2 And varlen = 4 Then
      If Val(varstr1) > 30 Then
         f_annee = f_remplit_gauche(varstr1, "19", varlen)
      Else
         f_annee = f_remplit_gauche(varstr1, "20", varlen)
      End If
   End If

End Function

Function f_basname(varstr1 As String) As String
'varstr1 adresse complète du fichier
'resultat = f_basname partie droite de l'adresse (fichier)
'           si rien après \ on ne renvoi rien
'           si aucun \ on renvoi la totalité

   Dim a$
   Dim varlig As String
   Dim varsepori As String
   Dim ccpt As Integer

   Dim tabbas() As String


   f_basname = varstr1

   varlig = varstr1

   varsepori = "\"

   ccpt = 0

plusreglebas:

   If InStr(varlig, varsepori) <> 0 Then
        ccpt = ccpt + 1
        ReDim Preserve tabbas(ccpt)
        tabbas(ccpt) = Left$(varlig, InStr(varlig, varsepori) - 1)
        varlig = Right$(varlig, Len(varlig) - InStr(varlig, varsepori) - Len(varsepori) + 1)
   Else
        GoTo finreglebas
   End If
        
   GoTo plusreglebas

finreglebas:

   ccpt = ccpt + 1
   ReDim Preserve tabbas(ccpt)
   tabbas(ccpt) = varlig

   If ccpt > 1 Then
      f_basname = tabbas(ccpt)
   End If


End Function

Function f_cara(varstr1 As String) As String
'test caractères
'varstr1 = texte


   Dim vartxt As String
   Dim varstr As String
   Dim varcar As String
   Dim i As Integer

   f_cara = varstr1
   
   vartxt = varstr1
   varstr = ""
   varcar = ""
   
   For i = 1 To Len(vartxt)
      varcar = Mid$(vartxt, i, 1)
      Select Case varcar
      Case "é"
         varcar = "e"
      Case "è"
         varcar = "e"
      Case "ç"
         varcar = "c"
      Case "à"
         varcar = "a"
      Case "-"
         varcar = " "
      Case "."
         varcar = " "
      Case ","
         varcar = " "
      Case ";"
         varcar = " "
      Case ":"
         varcar = " "
      Case "'"
         varcar = " "

      Case UCase$("é")
         varcar = "E"
      Case UCase$("è")
         varcar = "E"
      Case UCase$("ç")
         varcar = "C"
      Case UCase$("à")
         varcar = "A"
      Case UCase$("-")
         varcar = " "
      Case UCase$(".")
         varcar = " "
      Case UCase$(",")
         varcar = " "
      Case UCase$(";")
         varcar = " "
      Case UCase$(":")
         varcar = " "
      Case UCase$("'")
         varcar = " "
      Case Else
      End Select
      varstr = varstr & varcar
   Next i

   f_cara = varstr

End Function

Function f_champ(varstr1 As String, varstr2 As String, varvar1 As Variant) As String
'varstr1 = ligne a decouper
'varstr2 = separateur de champ
'varvar1 = position a recuperer   numéro champ:position ssv  varsepssv donne le séparateur ssv : par défaut = :
'resultat = f_champ champ numéro varpos

   Dim varval As String
   Dim varpos As Integer
   Dim varssv As Integer
   Dim vari As Integer
   Dim varj As Integer
   Dim varchp As String
   Dim varlen As Integer

   varval = ""
   
   If InStr(1, varvar1, ":", 1) > 0 Then
      varpos = Val(Left$(varvar1, InStr(1, varvar1, ":", 1) - 1))
      varssv = Val(Mid$(varvar1, InStr(1, varvar1, ":", 1) + 1))
   Else
      If IsNumeric(varvar1) Then
         varpos = Val(varvar1)
         varssv = 0
      Else
         MsgBox "ERREUR PARAMETRE f_champ : " & Str$(varvar1)
         Exit Function
      End If
   End If
   

   varj = 1
   varchp = varstr1
   'MsgBox varstr1
   varlen = Len(varstr2)
   For vari = 1 To varpos
   'MsgBox varchp
   'jld attention le séparateur doit être strictement identique
      varj = InStr(1, varchp, varstr2, 0)
      If varj > 0 Then
         On Error Resume Next
         varval = Left$(varchp, varj - 1)
         If Err Then
            f_champ = ""
            Err = 0
            Exit Function
         End If
         On Error GoTo 0

         varchp = Right$(varchp, Len(varchp) - varj - varlen + 1)
      Else
         varval = varchp
         varchp = ""
      End If
   Next vari

   'jld20020305 sous valeurs
   If varssv <> 0 Then
      If varsepssv = "" Then
         varsepssv = ":"
      End If
      varj = 1
      varchp = varval
      varlen = Len(varsepssv)
      For vari = 1 To varssv
         'jld attention le séparateur doit être strictement identique
         varj = InStr(1, varchp, varsepssv, 0)
         If varj > 0 Then
            On Error Resume Next
            varval = Left$(varchp, varj - 1)
            If Err Then
               f_champ = ""
               Err = 0
               Exit Function
            End If
            On Error GoTo 0
   
            varchp = Right$(varchp, Len(varchp) - varj - varlen + 1)
         Else
            varval = varchp
            varchp = ""
         End If
      Next vari
   End If

   f_champ = varval

End Function

Function f_champ1(varstr1 As String, varstr2 As String, varpos As Integer) As String
'varstr1 = ligne a decouper
'varstr2 = separateur de champ
'varpos = position a recuperer
'resultat = f_champ1 champ numéro varpos

   Dim vari As Integer
   Dim varj As Integer
   Dim varchp As String
   Dim varlen As Integer

   varlen = Len(varstr2)

   varj = 1
   varchp = varstr1
   For vari = 1 To varpos
      varj = InStr(1, varchp, varstr2)
      If varj > 0 Then
    On Error Resume Next
    f_champ1 = Left$(varchp, varj - 1)
    If Err Then
       f_champ1 = ""
       Err = 0
       Exit Function
    End If
    On Error GoTo 0

    varchp = Right$(varchp, Len(varchp) - varj - varlen + 1)
      Else
    f_champ1 = varchp
    varchp = ""
      End If
   Next vari

End Function

Function f_champs(varstr1 As String, varstr2 As String, varstr3 As String, varstr4 As String) As String
'permet de recomposer une ligne réduite à partir de sélection de champs
'varstr1 donnee a decouper : plusieurs champs
'varstr2 separateur de champ de donnee origine
'varstr3 liste des champs a recuperer separee par des varsep "," par defaut
'        exemple 1:1,3:2,9:3,7:4
'varstr4 separateur de champ de donnee destination
   
   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim varlig As String
   Dim varchp As String
   Dim varsepori As String
   Dim varsepdes As String
   Dim varchp1 As Integer
   Dim varchp2 As Integer
   Dim varmax As Integer
   Dim tabfin() As String
   ReDim tabfin(0)

   f_champs = ""

   varlig = varstr1
   varsepori = varstr2
   varchp = Trim$(varstr3)
   varsepdes = varstr4
   varchp1 = 0
   varchp2 = 0
   varmax = 0
   varstr = ""

   'sélection champ source et destination au format 1:1,2:2,...
   varnbr = 0
   varnbr = f_remplit_str(varchp, varsep)

   If varnbr > 0 Then
      For varint = 1 To varnbr
         varchp1 = 0
         varchp2 = 0
         varchp1 = Val(f_champ(tabstr(varint), ":", 1))
         varchp2 = Val(f_champ(tabstr(varint), ":", 2))
         'si champ destination non renseigné ou 0 : position par défaut varint
         If varchp2 = 0 Then varchp2 = varint
         
         If varchp2 > varmax Then
            varmax = varchp2
            ReDim Preserve tabfin(varmax)
         End If
         
         varrep = ""
         varrep = f_champ(varlig, varsepori, varchp1)
         tabfin(varchp2) = varrep

      Next varint
   
      For varint = 1 To UBound(tabfin)
         If varint = 1 Then
            'varstr = varrep
            varstr = tabfin(varint)
         Else
            'varstr = varstr & varsepdes & varrep
            varstr = varstr & varsepdes & tabfin(varint)
         End If
         
      Next varint

      f_champs = varstr
   
   End If

End Function

Function f_chp1(varstr1 As String) As Integer
'descri= recherche du champ cderepas pour un code
'varstr1= code
'retour= numéro de champ

   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim i As Integer

   f_chp1 = 0

   On Error Resume Next
   If UBound(tabdes1, 2) < 1 Then
      MsgBox "TABLEAU DE PARAMETRES des1 NON REMPLIT"
      Exit Function
   End If

   On Error GoTo 0

   varstr = Trim$(UCase$(varstr1))
   varrep = ""
   varint = 0

'cderepas,1,D,8,,,datcre,Date Saisie,

   For i = 1 To UBound(tabdes1, 2)
      varrep = Trim$(UCase$(tabdes1(7, i)))
      If varrep = varstr Then
         varint = Val(tabdes1(2, i))
         Exit For
      End If
   Next i

   f_chp1 = varint


End Function

Function f_chpsiz(varstr1 As String, varstr2 As String, varstr3 As String) As String
'permet de dimensionner la largeur des champs
'varstr1 donnee a formater : plusieurs champs
'varstr2 : separateur
'varstr3 : tailles des champs à redimensionner   'jld20010405
'tabchpsiz contient la largeur des champs

   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim varlig As String
   Dim varchp As String
   Dim varsepori As String
   Dim varsepdes As String
   Dim vartab As Integer
   Dim varchpsiz As String
   Dim i As Integer

   f_chpsiz = varstr1

   varlig = varstr1
   varsepori = varstr2
   varsepdes = varsepori
   varchpsiz = Trim$(varstr3)

   If varchpsiz = "" Then Exit Function

   varint = 0
   varint = f_remplit_str(varchpsiz, varsep)
   ReDim tabchpsiz(0)
   For i = 1 To varint
      ReDim Preserve tabchpsiz(i)
      tabchpsiz(i) = Val(tabstr(i))
   Next i

   On Error Resume Next
   vartab = UBound(tabchpsiz)
   If Err Then
      Err = 0
      Exit Function
   End If
   On Error GoTo 0


   varnbr = 0
   varnbr = f_remplit_str(varlig, varsepori)

   If varnbr > 0 Then
      varstr = ""
      
      For varint = 1 To varnbr
         varrep = tabstr(varint)
         If vartab >= varint Then
            If tabchpsiz(varint) > 0 Then
               varrep = f_remplit_droite(tabstr(varint), " ", tabchpsiz(varint))
            End If
         End If
         
         If varint = 1 Then
            varstr = varrep
         Else
            varstr = varstr & varsepdes & varrep
         End If
      Next varint
   
      'raz tabchpsiz
      For varint = 1 To vartab
          tabchpsiz(varint) = 0
      Next varint
      ReDim tabchpsiz(0)
   
      f_chpsiz = varstr
   
   End If


End Function

Function f_const() As String

   f_const = ""

   CrLf = Chr$(13) & Chr$(10)

   f_const = "OK"

End Function

Function f_coupe(varstr1 As String, varlon1 As Long) As Integer
'découpe une ligne en plusieurs morceaux : list = max 1024
'varstr1 = ligne à découper
'varlon1 = longueur
'retour = nombre de cellules dans tableau tabstr

   Dim varlig As String
   Dim varstr As String
   Dim varrep As String
   Dim varlon As Long
   Dim varcou As Long
   Dim vardeb As Long
   Dim varint As Integer
   Dim i As Long
   Dim j As Long

   f_coupe = 0

   ReDim tabstr(0)
   varstr = varstr1
   varcou = varlon1
   varlig = ""
   varlon = Len(varstr1)
   vardeb = 1
   varint = 0

   Do While varlon > varcou
       varrep = Mid$(varstr, vardeb, varcou)
       vardeb = vardeb + varcou
       varlon = varlon - varcou
       varint = varint + 1
       ReDim Preserve tabstr(varint)
       tabstr(varint) = varrep
      If Err Then 'jld20011203
         MsgBox Str(Err) & " ERREUR WHILE f_coupe"
         Err = 0
         Exit Do
      End If

   Loop
   
   If varlon <> 0 Then
      varrep = Mid$(varstr, vardeb, varcou)
      varint = varint + 1
      ReDim Preserve tabstr(varint)
      tabstr(varint) = varrep
   End If

   f_coupe = varint

End Function

Function f_cry(varstr1 As String, varstr2 As String) As String
'DESCRI=cryptage fichier
'varstr1=fichier
'varstr2=lecture ecriture R W
'retour dans le tableau tabcry
'si le fichier à lire n'est pas crypté : on le crypte

'iid.xxx
'secure
'lapreso
'chemin.lap
'lapchem
'alisutil.txt
'alislap.ini


   Dim varcod1 As Integer
   Dim varcod2 As Integer
   Dim varcar As String
   Dim varcho As String
   Dim varnom As String
   Dim varfic As String
   Dim varlig As String
   Dim varstr As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim nf As Integer
   Dim i As Integer
   Dim j As Integer
   Dim a$

   f_cry = ""

   varfic = Trim$(UCase$(varstr1))
   varcho = Trim$(UCase$(varstr2))

   If varfic = "" Or varcho = "" Then
      Exit Function
   End If

   varnom = Trim$(UCase$(f_champ(f_basname(varfic), ".", 1)))
   varnbr = 0

   If varcho = "R" Then

      On Error Resume Next
      nf = FreeFile
      Close #nf: Open varfic For Input As #nf
      If Err Then
         MsgBox Str(Err) & " ERREUR LECTURE fichier : " & varnom
         Close #nf
         Err = 0
         f_cry = ""
         Exit Function
      Else
         Select Case varnom
         Case "LAPRESO"
            ReDim tabcry(1, 0)
            Do While Not EOF(nf)
               'jld20040816
               On Error Resume Next
            
               Line Input #nf, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE"
                  Err = 0
                  Exit Do
               End If
               
               varlig = Trim$(a$)
               If varlig <> "" Then
                  varnbr = varnbr + 1
                  'test si crypté
                  If varnbr = 1 Then
                     If Asc(Left$(varlig, 1)) = 34 Then
                        'texte non crypté
                        varcho = "NON"
                        'jld20020105 on recrypte le fichier : parametre pour cryptage
                        'varcho = "W"
                     Else
                        varcho = "OUI"
                     End If
                     Err = 0 'jld20040816 err 5 si a$ vide : asc
                     
                  End If

                  If varcho = "OUI" Then
                     varstr = varlig
                     varlig = ""
                     For i = 1 To Len(varstr)
                        varcar = Mid$(varstr, i, 1)
                        varcod1 = Asc(varcar)
                        varcod2 = varcod1 Xor 255
                        varcar = Chr$(varcod2)
                        varlig = varlig & varcar
                     Next i
                     Err = 0 'jld20040816 err 5 si a$ vide : asc
                     
                  End If

                  varint = 0
                  varint = f_remplit_str(varlig, Chr$(34))
                  If varint = 3 Then
                     ReDim Preserve tabcry(1, varnbr)
                     tabcry(1, varnbr) = tabstr(2)
                  Else
                     'fichier non conforme
                     Close #nf
                     f_cry = ""
                     Exit Function
                  End If
               End If
            Loop
            If varnbr < 27 Then
               'fichier lapreso non conforme
               Close #nf
               f_cry = ""
               Exit Function
            End If

         Case Else
         End Select
      End If
      Close #nf

   End If
   If varcho = "W" Then
      
      Select Case varnom
      Case "LAPRESO"
            If UBound(tabcry, 2) < 27 Then
               MsgBox "ERREUR ECRITURE fichier : lapreso"
               f_cry = ""
               Exit Function

            End If
         
      Case Else
      End Select
      
      
      On Error Resume Next
      nf = FreeFile
      Close #nf: Open varfic For Output Lock Read Write As #nf  'jld20031203 : Lock #nf
      If Err Then
         MsgBox Str(Err) & " ERREUR ECRITURE fichier : " & varnom
         Err = 0
         f_cry = ""
         Exit Function
      Else
         Select Case varnom
         Case "LAPRESO"
         'attention au cryptage du guillemet
            For j = 1 To UBound(tabcry, 2)
               varlig = ""
               varstr = tabcry(1, j)
               For i = 1 To Len(varstr)
                  varcar = Mid$(varstr, i, 1)
                  varcod1 = Asc(varcar)
                  varcod2 = varcod1 Xor 255
                  varcar = Chr$(varcod2)
                  varlig = varlig & varcar
               Next i
               Err = 0 'jld20040816 err 5 si a$ vide : asc
               
               varlig = Chr$(34 Xor 255) & varlig & Chr$(34 Xor 255)
               Print #nf, varlig
            Next j
            
         Case Else
         End Select
      End If
      'jld20031224 Unlock #nf
      Close #nf

   End If

   If varcho <> "R" And varcho <> "W" Then
      f_cry = ""
      Exit Function
   End If

   f_cry = "OK"

End Function

Function f_cry2(varstr1 As String) As String
'DESCRI=cryptage ligne
'varstr1=ligne

   Dim varcod1 As Integer
   Dim varcod2 As Integer
   Dim varcar As String
   Dim varcho As String
   Dim varnom As String
   Dim varfic As String
   Dim varlig As String
   Dim varstr As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim nf As Integer
   Dim i As Integer
   Dim j As Integer
   Dim a$

   f_cry2 = varstr1

   varlig = varstr1

   If varlig = "" Then
      Exit Function
   End If

   varstr = varlig
   varlig = ""
   For i = 1 To Len(varstr)
      varcar = Mid$(varstr, i, 1)
      varcod1 = Asc(varcar)
      varcod2 = varcod1 Xor 255
      varcar = Chr$(varcod2)
      varlig = varlig & varcar
   Next i
   Err = 0 'jld20040816 err 5 si a$ vide : asc

   f_cry2 = varlig

End Function

Function f_ctol(varstr1 As Variant, varstr2 As String) As String
'DESCRI=Conversion prix chiffres en lettres (montant maxi : 999999.99)
'varstr1 = montant en chiffre ex:123.45 (2 décimales maxi)
'varstr2 = unité (Fr)
'retour = montant en lettres

   Dim prix1 As String
   Dim unite As String
   Dim resultat As String
   Dim ligne_cpl As String
   Dim prix_gauche As String
   Dim prix_droite As String
   Dim gauche_size As Integer
   Dim droite_size As Integer
   Dim already As Integer
   Dim isCentimes As Integer

     'on part du principe que le montant n'excede pas 999999.99
     'il faut donc connaitre la longueur du chiffre a gauche du point
     'on sait qu'à droite y'a toujours 2 chiffres

     prix1 = Trim$(varstr1)
     unite = Trim$(varstr2)

     prix_gauche = f_champ(prix1, ".", 1)
     prix_droite = f_champ(prix1, ".", 2)
     prix_droite = f_r_d(prix_droite, "0", 2)

    'on a séparé les deux parties

    '---------------------------------------------------
    'partie gauche :
     gauche_size = Len(prix_gauche)
     'y'a donc 6 tailles pour la gauche, on traite chacune séparément, si y'a plus que 6 (improbable)
     'alors on renvoie tout tel quel
            
            
     If gauche_size > 6 Then
            resultat = prix1
            GoTo finconv
     End If

     If gauche_size = 6 Then
        Select Case (Left(prix_gauche, 1))
            Case 1:
                 resultat = "CENT "
            Case 2:
                 resultat = "DEUX CENT "
            Case 3:
                 resultat = "TROIS CENT "
            Case 4:
                 resultat = "QUATRE CENT "
            Case 5:
                 resultat = "CINQ CENT "
            Case 6:
                 resultat = "SIX CENT "
            Case 7:
                 resultat = "SEPT CENT "
            Case 8:
                 resultat = "HUIT CENT "
            Case 9:
                 resultat = "NEUF CENT "
        End Select
            prix_gauche = Right(prix_gauche, 5)
            gauche_size = Len(prix_gauche)
     End If

     If gauche_size = 5 Then
        Select Case (Left(prix_gauche, 1))
            Case 1:
                 resultat = resultat & "DIX "
            Case 2:
                 resultat = resultat & "VINGT "
            Case 3:
                 resultat = resultat & "TRENTE "
            Case 4:
                 resultat = resultat & "QUARANTE "
            Case 5:
                 resultat = resultat & "CINQUANTE "
            Case 6:
                 resultat = resultat & "SOIXANTE "
            Case 7:
                 resultat = resultat & "SOIXANTE "
                 already = True
                 'attention, ici y'a une petite subtilité avec les 71,72,etc...
                    Select Case (Mid(prix_gauche, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX MILLE "
                        Case 1:
                             resultat = resultat & "ET ONZE MILLE "
                        Case 2:
                             resultat = resultat & "DOUZE MILLE "
                        Case 3:
                             resultat = resultat & "TREIZE MILLE "
                        Case 4:
                             resultat = resultat & "QUATORZE MILLE "
                        Case 5:
                             resultat = resultat & "QUINZE MILLE "
                        Case 6:
                             resultat = resultat & "SEIZE MILLE "
                        Case 7:
                             resultat = resultat & "DIX SEPT MILLE "
                        Case 8:
                             resultat = resultat & "DIX HUIT MILLE "
                        Case 9:
                             resultat = resultat & "DIX NEUF MILLE "
                    End Select
            Case 8:
                 resultat = resultat & "QUATRE VINGT "
            Case 9:
                 'attention, ici y'a une petite subtilité avec les 91,92,etc...
                 resultat = resultat & "QUATRE VINGT "
                 already = True
                    Select Case (Mid(prix_gauche, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX MILLE "
                        Case 1:
                             resultat = resultat & "ONZE MILLE "
                        Case 2:
                             resultat = resultat & "DOUZE MILLE "
                        Case 3:
                             resultat = resultat & "TREIZE MILLE "
                        Case 4:
                             resultat = resultat & "QUATORZE MILLE "
                        Case 5:
                             resultat = resultat & "QUINZE MILLE "
                        Case 6:
                             resultat = resultat & "SEIZE MILLE "
                        Case 7:
                             resultat = resultat & "DIX SEPT MILLE "
                        Case 8:
                             resultat = resultat & "DIX HUIT MILLE "
                        Case 9:
                             resultat = resultat & "DIX NEUF MILLE "
                    End Select
        End Select
            prix_gauche = Right(prix_gauche, 4)
            gauche_size = Len(prix_gauche)


     End If
     If gauche_size = 4 Then
        If already = False Then
            Select Case (Left(prix_gauche, 1))
                Case 0:
                     resultat = resultat & "MILLE "
                Case 1:
                     resultat = resultat & "ET UN MILLE "
                Case 2:
                     resultat = resultat & "DEUX MILLE "
                Case 3:
                     resultat = resultat & "TROIS MILLE "
                Case 4:
                     resultat = resultat & "QUATRE MILLE "
                Case 5:
                     resultat = resultat & "CINQ MILLE "
                Case 6:
                     resultat = resultat & "SIX MILLE "
                Case 7:
                     resultat = resultat & "SEPT MILLE "
                Case 8:
                     resultat = resultat & "HUIT MILLE "
                Case 9:
                     resultat = resultat & "NEUF MILLE "
            End Select
        End If
        already = False
        prix_gauche = Right(prix_gauche, 3)
        gauche_size = Len(prix_gauche)


     End If
     If gauche_size = 3 Then
        Select Case (Left(prix_gauche, 1))
            Case 1:
                 resultat = resultat & "CENT "
            Case 2:
                 resultat = resultat & "DEUX CENT "
            Case 3:
                 resultat = resultat & "TROIS CENT "
            Case 4:
                 resultat = resultat & "QUATRE CENT "
            Case 5:
                 resultat = resultat & "CINQ CENT "
            Case 6:
                 resultat = resultat & "SIX CENT "
            Case 7:
                 resultat = resultat & "SEPT CENT "
            Case 8:
                 resultat = resultat & "HUIT CENT "
            Case 9:
                 resultat = resultat & "NEUF CENT "
        End Select
            prix_gauche = Right(prix_gauche, 2)
            gauche_size = Len(prix_gauche)


     End If
     If gauche_size = 2 Then
        Select Case (Left(prix_gauche, 1))
            Case 1:
                 
                 already = True
                 'attention, ici y'a une petite subtilité avec les 11,12,etc...
                    Select Case (Mid(prix_gauche, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX "
                        Case 1:
                             resultat = resultat & "ONZE "
                        Case 2:
                             resultat = resultat & "DOUZE "
                        Case 3:
                             resultat = resultat & "TREIZE "
                        Case 4:
                             resultat = resultat & "QUATORZE "
                        Case 5:
                             resultat = resultat & "QUINZE "
                        Case 6:
                             resultat = resultat & "SEIZE "
                        Case 7:
                             resultat = resultat & "DIX SEPT "
                        Case 8:
                             resultat = resultat & "DIX HUIT "
                        Case 9:
                             resultat = resultat & "DIX NEUF "
                    End Select
                 
            Case 2:
                 resultat = resultat & "VINGT "
            Case 3:
                 resultat = resultat & "TRENTE "
            Case 4:
                 resultat = resultat & "QUARANTE "
            Case 5:
                 resultat = resultat & "CINQUANTE "
            Case 6:
                 resultat = resultat & "SOIXANTE "
            Case 7:
                 resultat = resultat & "SOIXANTE "
                 already = True
                 'attention, ici y'a une petite subtilité avec les 71,72,etc...
                    Select Case (Mid(prix_gauche, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX "
                        Case 1:
                             resultat = resultat & "ET ONZE "
                        Case 2:
                             resultat = resultat & "DOUZE "
                        Case 3:
                             resultat = resultat & "TREIZE "
                        Case 4:
                             resultat = resultat & "QUATORZE "
                        Case 5:
                             resultat = resultat & "QUINZE "
                        Case 6:
                             resultat = resultat & "SEIZE "
                        Case 7:
                             resultat = resultat & "DIX SEPT "
                        Case 8:
                             resultat = resultat & "DIX HUIT "
                        Case 9:
                             resultat = resultat & "DIX NEUF "
                    End Select
            Case 8:
                 resultat = resultat & "QUATRE VINGT "
            Case 9:
                 'attention, ici y'a une petite subtilité avec les 91,92,etc...
                 resultat = resultat & "QUATRE VINGT "
                 already = True
                    Select Case (Mid(prix_gauche, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX "
                        Case 1:
                             resultat = resultat & "ONZE "
                        Case 2:
                             resultat = resultat & "DOUZE "
                        Case 3:
                             resultat = resultat & "TREIZE "
                        Case 4:
                             resultat = resultat & "QUATORZE "
                        Case 5:
                             resultat = resultat & "QUINZE "
                        Case 6:
                             resultat = resultat & "SEIZE "
                        Case 7:
                             resultat = resultat & "DIX SEPT "
                        Case 8:
                             resultat = resultat & "DIX HUIT "
                        Case 9:
                             resultat = resultat & "DIX NEUF "
                    End Select
        End Select
            prix_gauche = Right(prix_gauche, 1)
            gauche_size = Len(prix_gauche)


     End If
     

     
     If gauche_size = 1 Then
        If already = False Then
            Select Case (Left(prix_gauche, 1))
            
                Case 1:
                     resultat = resultat & "ET UN "
                Case 2:
                     resultat = resultat & "DEUX "
                Case 3:
                     resultat = resultat & "TROIS "
                Case 4:
                     resultat = resultat & "QUATRE "
                Case 5:
                     resultat = resultat & "CINQ "
                Case 6:
                     resultat = resultat & "SIX "
                Case 7:
                     resultat = resultat & "SEPT "
                Case 8:
                     resultat = resultat & "HUIT "
                Case 9:
                     resultat = resultat & "NEUF "
            
            End Select
        End If

     End If

     already = False


    '---------------------------------------------------
    If Trim$(UCase$(resultat)) = "" Then
      resultat = "ZERO"
    End If
    If Left$(Trim$(UCase$(resultat)), 3) = "ET " Then
      resultat = Mid$(resultat, 4)
    End If
    If Trim$(unite) <> "" Then
      If Left$(Trim$(UCase$(resultat)), 1) = "U" Or Left$(Trim$(UCase$(resultat)), 1) = "Z" Then
         resultat = Trim$(resultat) & " " & unite
      Else
         resultat = Trim$(resultat) & " " & unite & "S"
      End If
    End If
    
    '---------------------------------------------------
    'partie droite :
    If Val(prix_droite) > 0 Then
        isCentimes = True
         resultat = resultat & " ET "
         droite_size = Len(prix_droite)
         If droite_size = 2 Then
            Select Case (Left(prix_droite, 1))
                Case 1:
                 already = True
                 'attention, ici y'a une petite subtilité avec les 11,12,etc...
                    Select Case (Mid(prix_droite, 2, 1))
                        Case 0:
                             resultat = resultat & "DIX "
                        Case 1:
                             resultat = resultat & "ONZE "
                        Case 2:
                             resultat = resultat & "DOUZE "
                        Case 3:
                             resultat = resultat & "TREIZE "
                        Case 4:
                             resultat = resultat & "QUATORZE "
                        Case 5:
                             resultat = resultat & "QUINZE "
                        Case 6:
                             resultat = resultat & "SEIZE "
                        Case 7:
                             resultat = resultat & "DIX SEPT "
                        Case 8:
                             resultat = resultat & "DIX HUIT "
                        Case 9:
                             resultat = resultat & "DIX NEUF "
                    End Select
                     
                Case 2:
                     resultat = resultat & "VINGT "
                Case 3:
                     resultat = resultat & "TRENTE "
                Case 4:
                     resultat = resultat & "QUARANTE "
                Case 5:
                     resultat = resultat & "CINQUANTE "
                Case 6:
                     resultat = resultat & "SOIXANTE "
                Case 7:
                     resultat = resultat & "SOIXANTE "
                     already = True
                     'attention, ici y'a une petite subtilité avec les 71,72,etc...
                        Select Case (Mid(prix_droite, 2, 1))
                            Case 0:
                                 resultat = resultat & "DIX "
                            Case 1:
                                 resultat = resultat & "ET ONZE "
                            Case 2:
                                 resultat = resultat & "DOUZE "
                            Case 3:
                                 resultat = resultat & "TREIZE "
                            Case 4:
                                 resultat = resultat & "QUATORZE "
                            Case 5:
                                 resultat = resultat & "QUINZE "
                            Case 6:
                                 resultat = resultat & "SEIZE "
                            Case 7:
                                 resultat = resultat & "DIX SEPT "
                            Case 8:
                                 resultat = resultat & "DIX HUIT "
                            Case 9:
                                 resultat = resultat & "DIX NEUF "
                        End Select
                Case 8:
                     resultat = resultat & "QUATRE VINGT "
                Case 9:
                     'attention, ici y'a une petite subtilité avec les 91,92,etc...
                     resultat = resultat & "QUATRE VINGT "
                     already = True
                        Select Case (Mid(prix_droite, 2, 1))
                            Case 0:
                                 resultat = resultat & "DIX "
                            Case 1:
                                 resultat = resultat & "ONZE "
                            Case 2:
                                 resultat = resultat & "DOUZE "
                            Case 3:
                                 resultat = resultat & "TREIZE "
                            Case 4:
                                 resultat = resultat & "QUATORZE "
                            Case 5:
                                 resultat = resultat & "QUINZE "
                            Case 6:
                                 resultat = resultat & "SEIZE "
                            Case 7:
                                 resultat = resultat & "DIX SEPT "
                            Case 8:
                                 resultat = resultat & "DIX HUIT "
                            Case 9:
                                 resultat = resultat & "DIX NEUF "
                        End Select
            End Select
                prix_droite = Right(prix_droite, 1)
                droite_size = Len(prix_droite)
    
    
         End If
     

     
         If droite_size = 1 Then
            If already = False Then
                Select Case (Left(prix_droite, 1))
                
                    Case 1:
                         resultat = resultat & "ET UN "
                    Case 2:
                         resultat = resultat & "DEUX "
                    Case 3:
                         resultat = resultat & "TROIS "
                    Case 4:
                         resultat = resultat & "QUATRE "
                    Case 5:
                         resultat = resultat & "CINQ "
                    Case 6:
                         resultat = resultat & "SIX "
                    Case 7:
                         resultat = resultat & "SEPT "
                    Case 8:
                         resultat = resultat & "HUIT "
                    Case 9:
                         resultat = resultat & "NEUF "
                
                End Select
            End If
    
         End If

    End If
    '---------------------------------------------------
    If isCentimes = True Then
      resultat = resultat & " CENTIMES"
      If Trim$(unite) <> "" Then
        If Left$(Trim$(UCase$(unite)), 1) = "A" Or Left$(Trim$(UCase$(unite)), 1) = "E" Or Left$(Trim$(UCase$(unite)), 1) = "I" Or Left$(Trim$(UCase$(unite)), 1) = "O" Or Left$(Trim$(UCase$(unite)), 1) = "U" Or Left$(Trim$(UCase$(unite)), 1) = "Y" Then
         resultat = resultat & " D'" & unite
        Else
         resultat = resultat & " DE " & unite
        End If
      End If
    End If
    ligne_cpl = "Arrête le présent Bordereau à la somme de " & resultat & "."

finconv:
     'f_ctol = ligne_cpl
     f_ctol = resultat

End Function

Function f_dadr(varstr1 As String, varstr2 As String, varstr3 As String) As String
'traitement des adresses par date ( accès rapide au répertoire )
'varstr1 = date 8 ex : 31122002 = 2002\20021230
'varstr2 = Adresse chemin avec AntéSlach ex : lapspe ou c:\spe\
'varstr3 = fichier ex : "fichier.txt"
'retour  = 2002\20021230\fichier.txt

   Dim varvar As Variant
   Dim vardat As String
   Dim varfic As String
   Dim vardir As String
   Dim varadr As String
   Dim varstr As String
   Dim varrep As String
   Dim varlon As Long
   Dim varint As Integer
   Dim varfin As String
   
   f_dadr = ""

   vardat = Trim$(varstr1)
   vardir = Trim$(varstr2)
   varfic = Trim$(varstr3)

   If vardat = "" Then Exit Function
   If vardir = "" Then Exit Function
   'If varfic = "" Then Exit Function

   varvar = vardat
   If Not IsNumeric(varvar) Then
      Exit Function
   End If

   vardat = f_temps(vardat, 3)
   If vardat = "" Then
      Exit Function
   End If

   If Right$(Trim$(vardir), 1) <> "\" Then
         vardir = vardir & "\"
   End If
   
   varadr = Left$(vardat, 4) & "\" & vardat & "\"
   varrep = f_mkdir(vardir & varadr, "", 0)
   If varrep <> "OK" Then
      Exit Function
   End If

   f_dadr = varadr & varfic

End Function

Function f_date(varstr1 As String, varstr2 As String, varlen As Integer) As String
'varstr1 = date jjmmyy
'varstr2 = separateur
'varlen = longueur annee
'resultat = f_date date formatée

   f_date = varstr1

   Select Case Len(varstr1)
      Case 6
         f_date = Left$(varstr1, 2) & varstr2 & Mid$(varstr1, 3, 2) & varstr2 & f_annee(Right$(varstr1, 2), varlen)

      Case 8
         f_date = Left$(varstr1, 2) & varstr2 & Mid$(varstr1, 3, 2) & varstr2 & Right$(varstr1, varlen)
      
      Case Else
   End Select

End Function

Function f_desmsq1(varstr1 As String) As String
'DESCRI=descrition des champs du masque : masque.des
'varstr1=masque

   Dim varmsq As String
   Dim varstr As String
   Dim varlig As String
   Dim varfic As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim varmax As Integer
   Dim nf As Integer
   Dim i As Integer
   Dim j As Integer
   Dim a$
   ReDim tabdes1(0, 0)

   f_desmsq1 = ""
   varmaxchp1 = 0       'nbre de champ dans le masque

   If lapdic = "" Then
      MsgBox "PARAMETRES WINLAP NON RENSEIGNES"
      Exit Function
   End If

   varmax = 0
   varnbr = 0
   varmsq = Trim$(UCase$(varstr1))
   'jld20020702
   If InStr(varmsq, ".") = 0 Then
      varfic = lapdic & varmsq & ".des"
   Else
      varfic = lapdic & varmsq
   End If

   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic For Input As #nf
   If Err Then
      Close #nf
      Err = 0
      Exit Function
   Else
      Do While Not EOF(nf)
         'jld20040110
         On Error Resume Next
         Line Input #nf, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE desmsq1"
            Err = 0
            Exit Do
         End If
         
         a$ = Trim$(a$)
         If a$ <> "" Then
            varint = 0
            varint = f_remplit_str(a$, varsep)
            If varint > 1 Then
               varnbr = varnbr + 1
               On Error Resume Next
               j = UBound(tabdes1, 1)
               If Err Or j = 0 Then
                  varmax = varint
                  ReDim tabdes1(varmax, 1)
               Else
                  If varmax <> varint Then
                     MsgBox "ERREUR DANS LA DESCRIPTION DU MASQUE : " & varmsq & "  " & Trim$(Str$(varmax)) & " <> " & Trim$(Str$(varint)) & " Champs"
                  End If
                  ReDim Preserve tabdes1(varmax, varnbr)
               End If
               
               For i = 1 To varint
                  tabdes1(i, varnbr) = tabstr(i)
               Next i
            End If
         End If
      Loop
   End If
   Close #nf
   On Error GoTo 0
   
   varmaxchp1 = varnbr
   f_desmsq1 = "OK"

End Function

Function f_dirname(varstr1 As String) As String
'varstr1 adresse complète du fichier
'resultat = f_dirname partie gauche de l'adresse (repertoire)
'           avec barre a la fin ex c:\tmp\
'           si aucun \ on ne renvoit rien

   Dim varstr As String
   Dim varint As Integer
   Dim varpos As Integer

   Dim a$
   Dim varlig As String
   Dim varsepori As String
   Dim ccpt As Integer

   Dim tabbas() As String


   f_dirname = varstr1

   varlig = varstr1

   varsepori = "\"

   ccpt = 0

plusregledir:

   If InStr(varlig, varsepori) <> 0 Then
        ccpt = ccpt + 1
        ReDim Preserve tabbas(ccpt)
        tabbas(ccpt) = Left$(varlig, InStr(varlig, varsepori) - 1)
        varlig = Right$(varlig, Len(varlig) - InStr(varlig, varsepori) - Len(varsepori) + 1)
   Else
        GoTo finregledir
   End If
        
   GoTo plusregledir

finregledir:

   ccpt = ccpt + 1
   ReDim Preserve tabbas(ccpt)
   tabbas(ccpt) = varlig

   varstr = ""
   For varpos = 1 To ccpt - 1
      varstr = varstr & tabbas(varpos) & varsepori
   Next

   f_dirname = varstr



End Function

Function f_dtoh(varstr1 As String) As String
'jld20030825 conversion décimal vers heure
'varstr1 = Décimal
'retour = format de sortie  heure

'1/4H     correspond à  0.25 en décimal
'1/2H     correspond à  0.5
'3/4H     correspond à 0.75
'1H         correspond à 1
'1H15     correspond à 1.25
'1H30     correspond à 1.50
'1H45     correspond à 1.75
'2H          correspond à 2


   Dim varstr As String
   Dim varheu As String
   Dim varmin As String
   Dim varres As Long
   Dim varlon As Long
   Dim varint As Integer
   Dim varvar As Variant
   Dim vartmp As String

   f_dtoh = "@@@"

   varstr = Trim$(UCase$(varstr1))
   varstr = f_remplace(varstr, ",", ".", "T", 1)
   varheu = "0"
   varmin = "0"

   'varsepdec
   'varsepwin
   
   varvar = varstr
   If Not IsNumeric(varvar) Then
      vartmp = f_remplace(varstr, ".", ",", "T", 1)
      varvar = vartmp
      If Not IsNumeric(varvar) Then
         MsgBox ("FORMAT DONNEE NON NUMERIQUE")
         Exit Function
      End If
   End If

   varheu = Trim$(f_champ(varstr, ".", 1))
   varmin = Left$(Trim$(f_champ(varstr, ".", 2)), 2)
   varmin = f_r_d(Trim$(f_champ(varstr, ".", 2)), "0", 2)

   varres = Val(varmin)
   
   If varres >= 100 Then
      MsgBox "ERREUR FORMAT DECIMAL (MINUTES) : " & varmin
      Exit Function
   End If
   
   
   varlon = Int((varres / 100) * 60)
   If (((varres / 100) * 60) - varlon) >= 0.5 Then
      varres = varlon + 1
   Else
      varres = varlon
   End If

   varmin = f_r_g(Trim$(Str$(varres)), "0", 2)


   f_dtoh = varheu & "H" & varmin


End Function

Function f_existe_regle(varstr1 As String, varstr2 As String) As String
'descri= Recherche règle dans le fichier de règles du masque (tableau)
'un pointeur static varposreg permet de mémoriser la position dans le tableau
'varstr1 = nom de la règle
'varstr2 = D : debut, S : suite
'          varposreg donne la position dans le tableau de règles
'          D réinitialise la position à 1
'          S lit la règle suivante dans le tableau
'jldyyy      revoir le système vardeb D S : obsolète avec variable varposreg
'retour = ligne de la règle
'retour = tabstr avec contenu de la règle

   Dim reglegauche As String
   Dim varstr As String
   Dim varreg As String
   Dim vardeb As String
   Dim varpos As Integer
   Dim varint As Integer
   Dim i As Long

   f_existe_regle = ""

   varreg = Trim$(UCase$(varstr1))
   vardeb = Left$(Trim$(UCase$(varstr2)), 1)

   If vardeb = "D" Then
      varposreg = 1
   End If

   For i = varposreg To nbregle
      varposreg = i + 1
      If InStr(regle(i), ",") > 0 Then
         reglegauche = Trim$(UCase$(f_champ(regle(i), varsep, 1)))
         If reglegauche = varreg Then
            varint = 0
            varint = f_remplit_str(regle(i), varsep)
            If varint > 1 Then
               f_existe_regle = regle(i)
               Exit For
            End If
         End If
      End If
   Next i

End Function

 Function f_format(varstr1 As String, varstr2 As String) As String
'varstr1 = donnée à formater
'varstr2 = format

   Dim varstr As String
   Dim vardon As String
   Dim varfor As String
   Dim vargau As String
   Dim vardro As String
   Dim varnbrdec As Integer
   Dim varsepmil As Integer
   Dim varint As Integer
   Dim i As Integer
   Dim j As Integer

   vardon = Trim$(varstr1)
   varfor = Trim$(varstr2)
   varnbrdec = 0
   varsepmil = 0

   f_format = vardon


   'analyse du format
   'ex = # ###.##

   'recherche du caractère #  : les autres formats ensuite
   If InStr(1, varfor, "#", 1) = 0 Then
         Screen.MousePointer = 0
         Exit Function
   End If

   'recherche point
   If InStr(1, varfor, ".", 1) = 0 Then
      varnbrdec = 0
      vardro = ""
      vargau = f_champ(vardon, ".", 1)
   Else
      varnbrdec = Len(f_champ(varfor, ".", 2))
      vardro = Left$(f_champ(vardon, ".", 2) & String$(varnbrdec, "0"), varnbrdec)
      vargau = f_champ(vardon, ".", 1)
   End If

   'recherche sep mil
   If InStr(1, varfor, "# #", 1) = 0 Then
      varsepmil = 0
   Else
      varsepmil = 1
      varstr = f_champ(varfor, ".", 1)
      varint = 0
      varint = f_remplit_str(varstr, " ")
      varsepmil = Len(tabstr(varint))
      varstr = ""
      j = 0
      If varsepmil > 0 Then
         For i = Len(vargau) To 1 Step -1
            varstr = Mid$(vargau, i, 1) & varstr
            j = j + 1
            If j = varsepmil Then
               j = 0
               varstr = " " & varstr
            End If
         Next i
         vargau = Trim$(varstr)
      End If
   End If

   'formatage
   
   If varnbrdec = 0 Then
      f_format = vargau
   Else
      f_format = vargau & "." & vardro
   End If


End Function

Function f_fusion(varnum As Variant) As String

'Cette fonction piste toutes les informations de fusions
'incluses dans un num, et renvoie le num actuel du dossier
'varnum = num d'origine
'f_fusion = num actuel du dossier

   Dim varvar As Variant
   Dim varstr As String
   Dim faitfus As Integer
   Dim contfus As Long
   Dim lect As String
   Dim numfus As String
   Dim tmpfus As String
   Dim nf As Integer
   Dim varmsg As String
   Dim vardos As String

   'jld20021120
   'f_fusion = varnum

   vardos = CStr(varnum)

   varmsg = Trim$(f_champ(vardos, ":", 2))
   vardos = Trim$(f_champ(vardos, ":", 1))
   If varmsg = "" Then
      varmsg = "1"
   End If

   f_fusion = vardos


   contfus = 1
   numfus = vardos

   'jld20030226
   If Trim$(numfus) = "" Then Exit Function
   varvar = numfus
   If Not IsNumeric(varvar) Then Exit Function

   Do While contfus = 1

      faitfus = 0
      'jld20021120
      'jld20040205 jmd signale des erreurs 70 en accès répeté à partir de l agenda : suppression du lock
      nf = FreeFile
      On Error Resume Next
      'Close #nf: Open lapnum + numfus For Input Lock Read Write As #nf ':lock #nf
      Close #nf: Open lapnum + numfus For Input As #nf
      If Err = 0 Then
         Do While Not EOF(nf)
            Line Input #nf, lect
            If InStr(lect, "DOSSIER FUSION VERS : ") Then
               tmpfus = Trim$(f_champ(lect, " : ", 2))
               If tmpfus = "" Then
                  f_fusion = numfus
                  'Unlock #nf
                  'Close #nf
                  'Exit Function
                  'FRED20021118
                  
                  GoTo fusfin

               Else
                  numfus = tmpfus
               End If
               faitfus = 1
               f_fusion = numfus
               If numfus = vardos Then 'On tourne en rond, dossier fusionne deux fois vers le meme numero, ca boucle
                  'Unlock #nf
                  'Close #nf
                  'Exit Function
                  'FRED20021118

                  GoTo fusfin

               End If
               If Err Then
                  f_fusion = numfus
                  contfus = 0
               End If
               Exit Do
            End If
         Loop
         
         If faitfus = 0 Then
            f_fusion = numfus
            contfus = 0
            Exit Do
         End If
      Else
         'jld20030226 pas de message : pb dans agendas : jmd
         If varmsg = "1" Then
            'jld20030331
            'jld20040209 ajout erreur 75 = fichier inexistant avec vb6
            If Err <> 53 And Err <> 52 And Err <> 75 Then
               MsgBox Str(Err) & " Erreur lecture NUM dossier : " + vardos
            End If
         End If
         GoTo fusfin
      End If
      'jld20031224 Unlock #nf
      Close #nf

   Loop

fusfin:

   'jld20031224 Unlock #nf
   Close #nf

End Function

Function f_htod(varstr1 As String) As String
'jld20030825 conversion heure vers décimal
'varstr1 = heure
'retour = format de sortie Décimal

'1/4H     correspond à  0.25 en décimal
'1/2H     correspond à  0.5
'3/4H     correspond à 0.75
'1H         correspond à 1
'1H15     correspond à 1.25
'1H30     correspond à 1.50
'1H45     correspond à 1.75
'2H          correspond à 2


   Dim varstr As String
   Dim varheu As String
   Dim varmin As String
   Dim varres As Long
   Dim varlon As Long
   Dim varint As Integer

   f_htod = "@@@"

   varstr = Trim$(UCase$(varstr1))
   varheu = "0"
   varmin = "0"

   'détermination du format des données
   varstr = f_remplace(varstr, " ", "", "T", 1)

   varint = f_remplit_str(varstr, "H")

   If varint > 1 Then
      Select Case tabstr(1)
      Case "1/4"
         varmin = "15"
         
      Case "2/4"
         varmin = "30"

      Case "1/2"
         varmin = "30"

      Case "3/4"
         varmin = "45"

      Case Else
         varheu = Trim$(tabstr(1))
         varmin = Trim$(tabstr(2))

      End Select

      If Val(varmin) >= 60 Then
         MsgBox "ERREUR FORMAT MINUTES : " & varmin
         Exit Function
      End If

      'varmin = f_r_g(Trim$(Str$(Int((Val(varmin) * 100) / 60))), "0", 2)
      
      varres = Val(varmin)
      varlon = Int((varres / 60) * 100)
      If (((varres / 60) * 100) - varlon) >= 0.5 Then
         varres = varlon + 1
      Else
         varres = varlon
      End If

      varmin = f_r_g(Trim$(Str$(varres)), "0", 2)
      
      f_htod = varheu & "." & varmin

   Else

      f_htod = varstr
      
   End If

End Function

Function f_laptra(varstr1 As String, varstr2 As String, varstr3 As String, varstr4 As String, varstr5 As String, varstr6 As String, varstr7 As String, varstr8 As String) As String
'trace des actions classées par date YYYY\YYYYMM
'varstr1 = masque
'varstr2 = fichier
'varstr3 = libre
'varstr4 = adresse
'varstr5 = action
'varstr6 = erreur
'varstr7 = libre
'varstr8 = ligne avec séparateur point virgule
'
'   varstr = f_laptra(varmsq, varobjide, varobjseq, varobjloc, varaction, varerr, varlibre, varlig)


'date
'heure
'libre
'userid
'service
'libre
'libre
'compteur
'iupa

   Dim vardir As String
   Dim vardatinv As String
   Dim varerr As Integer
   Dim varlon As Long
   Dim varnumlog As String
   Dim varuserid As String
   Dim varser As String
   Dim vardat As String
   Dim varheu As String
   Dim vardos As String
   Dim varnature As String
   Dim varlib As String
   Dim varmsq As String
   Dim varchpnat As Integer
   Dim varchpide As Integer
   Dim varchpseq As Integer
   Dim varchploc As Integer
   Dim varobjnat As String
   Dim varobjide As String
   Dim varobjseq As String
   Dim varobjloc As String
   Dim varaction As String
   Dim varstatus As String
   Dim varligdoc As String
   Dim varinf As String
   Dim varstr As String
   Dim varfic As String
   Dim varlig As String
   Dim varint As Integer
   Dim nf As Integer
   Dim nf2 As Integer
   Dim i As Integer
   Dim j As Integer
   Dim a$
   Dim vartabdoc() As String
   ReDim vartabdoc(10, 0)

   f_laptra = ""

   varmsq = Trim$(UCase$(varstr1))
   varobjide = Trim$(varstr2)
   varobjseq = Trim$(varstr3)
   varobjloc = Trim$(varstr4)
   varaction = Trim$(varstr5)
   varstatus = Trim$(varstr6)
   varnature = Trim$(UCase$(varstr7))
   varligdoc = Trim$(varstr8)
   'varligdoc = f_remplace(varligdoc, ",", ";", "T", 1)

   varuserid = lapide
   varser = lapser
   vardat = Format(Now, "DDMMYYYY")
   varheu = Time$
   vardos = lapdos
   

   'stockage du fichier de log dans le répertoire LAPLOG

      varerr = 0
      varstr = ""
      On Error Resume Next
      nf2 = FreeFile
   
      varlig = vardat
      varlig = varlig & varsep & varheu
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & varuserid
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & varser
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & vardos
      varlig = varlig & varsep & varmsq
      varlig = varlig & varsep & varobjide
      varlig = varlig & varsep & varobjseq
      varlig = varlig & varsep & varobjloc
      varlig = varlig & varsep & varaction
      varlig = varlig & varsep & varstatus
      varlig = varlig & varsep & varobjnat
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & ""
      varlig = varlig & varsep & varligdoc
      
      vardir = laptra
      vardir = vardir & Right$(vardat, 4) & "\"
      vardir = vardir & Left$(f_temps(vardat, 3), 6) & "\"
      On Error Resume Next
      varstr = f_mkdir(vardir, "", 0)
      
      If varstr = "OK" Then
   
         varerr = 0
         On Error Resume Next
         nf = FreeFile
         Close #nf: Open vardir & "LAPTRA.STD" For Append Lock Read Write As #nf  'jld20031203 : Lock #nf
         If Err Then
            varerr = Err
            MsgBox Str(Err) & " ERREUR ECRITURE fichier log : LAPTRA.STD"
         Else
               Print #nf, "LAPTRA"; varsep; "¥"; Date$; varsep; varheu; varsep; varlig
         End If
         'jld20031224 Unlock #nf
         Close #nf
      
      End If

      If varerr = 0 Then
         f_laptra = "OK"
      End If

End Function

Function f_lapvar(varstr1 As String) As String
'varstr1 = fichier à lire : par defaut c:\vb\lapreso

' 1 "c:\med\"
' 2 "c:\par\"       "admsej"
' 3 "c:\ant\"       "jrs"
' 4 "c:\dia\"       "c:\repas\"
' 5 "c:\med\"       "histo cde c:\repash\ : laprepsv lapregrg"
' 6 "c:\paradm\"    "dic repas hopale"
' 7 "c:\parmed\"    "documents externes   : H lapdoc dip hopale"
' 8 "c:\ent\"       "repertoire med repas : .ver"
' 9 "c:\aig\"       "documents externes   : H lapdoc .doc"
'10 "c:\aud\"       "documents images JPG : I lapimg"
'11 "c:\dad\"       "documents DICOM      : D lapimg"
'12 "c:\res\"       "documents son        : S"
'13 "c:\ree\"       "documents video      : V"
'14 "c:\rea\"
'15 "c:\rad\"
'16 "c:\rud\"
'17 "c:\ndi\"
'18 "c:\DIC\"
'19 "c:\num\"
'20 "c:\vb\"
'21 "c:\"
'22 "c:\dic\"
'23 "c:\pha\"
'24 "c:\trt\"
'25 "c:\lap\"
'26 "c:\spe\"
'27 "c:\"
'28
'29 "c:\rapport"
         
   Dim sss0 As Long
   Dim sss1 As Long
   
   Dim varasc As Integer
   Dim varstr As String
   Dim varrep As String
   Dim varfic As String
   Dim nf2 As Integer
   Dim nf As Integer
   Dim t As Integer
   Dim a$

   f_lapvar = ""

   laproot = ""
   laphlp = ""
   lapdic = ""
   laplap = ""
   lapmed = ""
   lapnum = ""
   lapvb = ""
   lappro = "" 'jld20031207
   lapdi2 = ""
   lapspe = ""
   laprap = ""
   lapsur = ""
   laptmp = ""
   'jld20040909
   laprem = ""
   
   lapage = ""
   lapnom = ""
   lappre = ""
   lapnai = ""
   lapsex = ""
   lapdos = ""
   lapnna = ""
   lapdoc = ""
   
   lapser = ""
   laprep = ""
   lapreph = ""
   lapdich = ""
   lapleth = ""
   laptrt = ""
   
   'lappro = "" 'dernier programme utilisé
   lapdat = "" 'date en cours
   lapheu = "" 'heure en cours

   lapdat = Date$
   lapheu = Time$

   'fred20030429
   lapana = ""
   lapbac = ""

   varstr1 = Trim$(varstr1)
   If varstr1 = "" Then
      varfic = "c:\vb\lapreso"
   Else
      varfic = varstr1
   End If
   
'GoTo suitelapvar

   'jld20021217 élimination des "
   varfic = f_remplace(varfic, """", "", "T", 1)

    nf = FreeFile
    On Error Resume Next
    Close #nf: Open varfic For Input As #nf
      If Err Then
         MsgBox Str$(Err) & " ERREUR LEC : " & varfic 'f_basname(varfic)
         Close #nf
         Err = 0
         Exit Function
      End If

      'recopie lapreso pour autres applis si c:\vb existe
      'On Error Resume Next
      'MkDir "c:\vb"
      
      'jld20010321
      'On Error Resume Next
      'FileCopy varfic, "c:\vb\lapreso"
      'Err = 0

      t = 0
      varstr = ""

      Do While Not EOF(nf)
         
         'Input #nf, varstr     err
         'jld20040816
         On Error Resume Next
         
         Line Input #nf, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE lapvar"
            Err = 0
            Exit Do
         End If

         'jld20010925
         varasc = Asc(Left$(Trim$(a$), 1))
         Err = 0 'jld20040816 err 5 si a$ vide : asc
         
            If (varasc And 128) = 128 Then
               a$ = f_cry2(a$)
            End If

         a$ = f_remplace(a$, """", "", "T", 1)
         varstr = Trim$(UCase$(f_champ(a$, varsep, 1)))

         t = t + 1
         If t <= 30 Then
            di$(t) = Trim$(varstr)
         Else
            varstr = Trim$(varstr)
            MsgBox ("TROP DE LIGNES : " & Str(t) & " = " & varstr)
         End If

      Loop
         
      Close #nf
      On Error GoTo 0

'suitelapvar:

      lapmed = Trim$(UCase$(di$(1)))   'med ou adm
      lapbac = Trim$(UCase$(di$(2)))   'fred20030429 itmtx adresse bactériologie
      lapana = Trim$(UCase$(di$(3)))   'fred20030429 itmtx adresse anatomopathologie anapath
      
      laprep = Trim$(UCase$(di$(4)))   'repas hopale : voir pour adresse 4
      lapreph = Trim$(UCase$(di$(5)))  'repas histo cde : voir pour adresse 5
      lapdich = Trim$(UCase$(di$(6)))  'dic repas
      lapleth = Trim$(UCase$(di$(7)))  'lettres hors winlap : dip
      laphlp = Trim$(UCase$(di$(8)))  'doc
      lapres = Trim$(UCase$(di$(12)))
      lapdic = Trim$(UCase$(di$(18)))  'dic
      lapnum = Trim$(UCase$(di$(19)))  'num
      lapvb = Trim$(UCase$(di$(20)))   'vb
      'jld20031207 jldyyy
      If vbreseau <> "" Then
         lappro = vbreseau
      Else
         lappro = lapvb
      End If
      laproot = Trim$(UCase$(di$(21))) 'root
      lapdi2 = Trim$(UCase$(di$(22)))  'dic22
      laptrt = Trim$(UCase$(di$(24)))  'genebio ?
      laplap = Trim$(UCase$(di$(25)))  'lap
      lapspe = Trim$(UCase$(di$(26)))  'specif : agendas
      laplet = Trim$(UCase$(di$(27)))  'lettres
      laprap = Trim$(UCase$(di$(29)))  'modèles rapport
      lapsur = Trim$(UCase$(di$(30)))

      If laproot = "" Then
         MsgBox ("ERREUR LEC variables LAP : " & f_basname(varfic))
         Exit Function
      End If


      If Len(lapmed) > 0 And Right$(lapmed, 1) <> "\" Then lapmed = lapmed & "\"
      'fred20030429
      If Len(lapbac) > 0 And Right$(lapbac, 1) <> "\" Then lapbac = lapbac & "\"
      If Len(lapana) > 0 And Right$(lapana, 1) <> "\" Then lapana = lapana & "\"
      If Len(laprep) > 0 And Right$(laprep, 1) <> "\" Then laprep = laprep & "\"
      If Len(lapreph) > 0 And Right$(lapreph, 1) <> "\" Then lapreph = lapreph & "\"
      If Len(lapdich) > 0 And Right$(lapdich, 1) <> "\" Then lapdich = lapdich & "\"
      If Len(lapleth) > 0 And Right$(lapleth, 1) <> "\" Then lapleth = lapleth & "\"
      If Len(laphlp) > 0 And Right$(laphlp, 1) <> "\" Then laphlp = laphlp & "\"
      If Len(lapres) > 0 And Right$(lapres, 1) <> "\" Then lapres = lapres & "\"
      If Len(lapdic) > 0 And Right$(lapdic, 1) <> "\" Then lapdic = lapdic & "\"
      If Len(lapnum) > 0 And Right$(lapnum, 1) <> "\" Then lapnum = lapnum & "\"
      If Len(lapvb) > 0 And Right$(lapvb, 1) <> "\" Then lapvb = lapvb & "\"
      'jld20031207
      If Len(lappro) > 0 And Right$(lappro, 1) <> "\" Then lappro = lappro & "\"
      If Len(laproot) > 0 And Right$(laproot, 1) <> "\" Then laproot = laproot & "\"
      If Len(lapdi2) > 0 And Right$(lapdi2, 1) <> "\" Then lapdi2 = lapdi2 & "\"
      If Len(laptrt) > 0 And Right$(laptrt, 1) <> "\" Then laptrt = laptrt & "\"
      If Len(laplap) > 0 And Right$(laplap, 1) <> "\" Then laplap = laplap & "\"
      If Len(lapspe) > 0 And Right$(lapspe, 1) <> "\" Then lapspe = lapspe & "\"
      If Len(laplet) > 0 And Right$(laplet, 1) <> "\" Then laplet = laplet & "\"
      If Len(laprap) > 0 And Right$(laprap, 1) <> "\" Then laprap = laprap & "\"
      If Len(lapsur) > 0 And Right$(lapsur, 1) <> "\" Then lapsur = lapsur & "\"



      laptmp = laproot & "LAPTMP"
      On Error Resume Next
      varstr = ""
      varstr = Dir(laptmp, 16)
      If varstr = "" Then
         MkDir laptmp
      End If
      On Error GoTo 0
      laptmp = laproot & "LAPTMP\"

      'jld20040909
      laprem = laptmp & "REPREM"
      On Error Resume Next
      varstr = ""
      varstr = Dir(laprem, 16)
      If varstr = "" Then
         MkDir laprem
      End If
      On Error GoTo 0
      laprem = laptmp & "REPREM\"

      'jld20031023 trace au meme niveau que med

      'laptra = laproot & "LAPLOG"
      
      Dim varint As Integer
      varint = Len(lapmed)
      If varint > 1 Then
         varstr = Left$(lapmed, varint - 1)
         varstr = f_dirname(varstr)
         laptra = varstr & "LAPLOG"
      Else
         laptra = laproot & "LAPLOG"
      End If
      
      On Error Resume Next
      varstr = ""
      varstr = Dir(laptra, 16)
      If varstr = "" Then
         MkDir laptra
      End If
      On Error GoTo 0
      laptra = laptra & "\"


      'si SERVICE les modèles seront dans le repertoire lettres + service + MODELE
      If Right$(Trim$(UCase$(laprap)), 9) <> "\RAPPORT\" And Right$(Trim$(UCase$(laprap)), 9) <> "\SERVICE\" Then
         laprap = laproot & "RAPPORT\"
      End If

      If Right$(Trim$(UCase$(laprep)), 7) <> "\REPAS\" Then
         laprep = laproot & "REPAS\"
      End If

      If Right$(Trim$(UCase$(lapreph)), 8) <> "\REPASH\" Then
         lapreph = laproot & "REPASH\"
      End If

      If Trim$(lapdich) = "" Then
         lapdich = lapdic
      End If
      
      If Trim$(lapdi2) = "" Then
         lapdi2 = lapdic
      End If

      If Trim$(laphlp) = "" Then
         laphlp = laproot & "LAPHLP\"
      End If




      varlaphlp = laphlp
      varlapdic = lapdic
      varlaplap = laplap
      varlapmed = lapmed
      varlapnum = lapnum
      varlapvb = lapvb
      'jld20031207
      varlappro = lappro
      varlaproot = laproot
      varlapdi2 = lapdi2
      varlapres = lapres
      varlapspe = lapspe
      varlaplet = laplet
      varlapleth = lapleth
      varlapsur = lapsur
      varlaprap = laprap
      varLaprep = laprep
      varLapreph = lapreph
      varLapdich = lapdich
      varlaptrt = laptrt
      varlaptmp = laptmp

      'jld20011107 identité utilisateur
      'login.txt champ 3
      lapide = ""
      'jld20030512
      lapmet = ""
      lapniv = ""
      varfic = lapvb & "login.txt"
      nf = FreeFile
      On Error Resume Next
      Close #nf: Open varfic For Input As #nf
      If Err Then
         If Err = 53 Then 'jld20011107
            Close #nf
            nf = FreeFile
            On Error Resume Next
            Close #nf: Open varfic For Output As #nf
               Print #nf, Format$(Now, "DDMMYYYY"); varsep; Time$; varsep; "***"; varsep; varsep
            Close #nf
            Err = 0
         Else
            MsgBox Str$(Err) & " ERREUR LEC fichier : " & f_basname(varfic)
            lapide = "***"
            'Exit Function
         End If
         Err = 0
      Else
         Line Input #nf, varstr
         lapide = Trim$(UCase$(f_champ(varstr, varsep, 3)))
         'jld20030512
         lapmet = Trim$(UCase$(f_champ(varstr, varsep, 4)))
         lapniv = Trim$(UCase$(f_champ(varstr, varsep, 5)))
      End If
      Close #nf
      On Error GoTo 0
      If lapide = "" Then lapide = "****"
      'jld20030512
      If lapmet = "" Then lapmet = "TEST"
      If lapniv = "" Then lapniv = "1"

      
'jld20021118 voir utilité de lire essai2 au démarrage
GoTo suiteessai2

      'jld20021118 aider
      If lapmmi = 2 Then
         varfic = laproot & "numerow"
      Else
         'jld20040413 masqnorm multienvironnement
         'varfic = laproot & "essai2"
         varfic = laproot & lapmsq & "essai2"
      End If

         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf
      If Err Then
         MsgBox Str$(Err) & " ERREUR LEC fichier : " & f_basname(varfic)
         Err = 0
         lapdos = ""
         Exit Function
      Else
         Line Input #nf, varstr
         lapdos = Trim$(f_champ(varstr, varsep, 1))
      End If
      Close #nf
      On Error GoTo 0

      lapage = ""
      lapnom = ""
      lappre = ""
      lapnai = ""
      lapsex = ""
      lapnna = ""

      If lapdos <> "" Then
         'jld20020805 millier
         Select Case lapnmi
         Case 1
            lappat = f_1000(lapdos, lapnum)
         Case Else
            lappat = lapdos
         End Select

         'varfic = lapnum & lapdos
         varfic = lapnum & lappat

         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf 'jld20031203 : Lock #nf
         If Err = 0 Then
            Line Input #nf, lapnom
            Line Input #nf, lappre
            Line Input #nf, lapnai
            Line Input #nf, lapsex
            On Error Resume Next
            Line Input #nf, varstr
            On Error Resume Next
            Line Input #nf, lapnna
         Else
            Err = 0
         End If
         'jld20031203 Unlock #nf
         Close #nf
         On Error GoTo 0

         Err = 0
         lapnom = ascii_ansi(Trim$(UCase$(lapnom)))
         lappre = ascii_ansi(Trim$(UCase$(lappre)))
         lapnai = ascii_ansi(Trim$(f_temps(lapnai, 2)))
         lapsex = ascii_ansi(Trim$(UCase$(lapsex)))
         varstr = ascii_ansi(Trim$(varstr))
         lapnna = ascii_ansi(Trim$(UCase$(lapnna)))
         
         If Trim$(varstr) <> "" And Trim$(f_champ(varstr, " ", 1)) <> lapdos Then
            MsgBox "ERREUR NUMERO DOSSIER dans le fichier NUM : " & lapdos & "  " & varstr
         End If
      End If

      
      'FRED20030321
      'lapage = Trim$(Str$((Val(f_temps("01" + Left$(Date$, 2) + Right$(Date$, 4), 15)) - Val(f_temps(lapnai, 15))) \ 365))
      If globnote <> "" And lapapp = "AGENDACB" Then
         lapage = Trim$(Str$(Int((Val(f_temps(f_temps(globnote, 2), 15)) - Val(f_temps(lapnai, 15))) / 365.2)))
      Else
         lapage = Trim$(Str$(Int((Val(f_temps(Mid$(Date$, 4, 2) + Left$(Date$, 2) + Right$(Date$, 4), 15)) - Val(f_temps(lapnai, 15))) / 365.2)))
      End If
      If Val(lapage) < 0 Or Val(lapage) > 150 Then
         lapage = ""
      End If

suiteessai2:

      'jld20040413 masqnorm multienvironnement
      'varfic = laproot & "se"
      varfic = laproot & lapmsq & "se"
         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf
      If Err Then
         MsgBox Str$(Err) & " ERREUR LEC : " & f_basname(varfic)
         Err = 0
         lapser = ""
         Exit Function
      Else
         Line Input #nf, varstr
         lapser = Trim$(varstr)
      End If
      Close #nf
      On Error GoTo 0


      f_lapvar = "OK"

End Function

Function f_lecalislap() As String
'reservé à win_lap : pas de gestion du parametre MULTIBASE
    
    Dim varwindir As String
    Dim ficreso As String
    Dim varint As Integer
    Dim varstr As String
    Dim varnbrbas As Integer
    Dim varnbrser As Integer
    Dim find As Integer
    Dim find2 As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim nf As Integer
    Dim varrep As String
    Dim varlig As String
    Dim varfic As String
    Dim a$
    Dim b$
    Dim cr As String
    Dim varalis As String
    Dim varser As String
    ReDim tabser(0)

   'jld20011107 gestion doc
    Dim varnbr As Integer
    ReDim tabdocser(0)
    ReDim tabdocpra(0)
    ReDim tabdocmod(0)
    ReDim tabdocide(0)


    Dim varutil As String


   f_lecalislap = "c:\vb\lapreso"

    varstr = ""
    'attention argument pour lancement direct masque
    'jld20010323 varutil = Trim$(laparg)
    If varutil = "" Then varutil = "alisutil.txt"

   cr = Chr$(13) & Chr$(10)
   varwindir = ""
   varstr = ""
   ficreso = ""
   varint = 0
   varnbrbas = 0
   varnbrser = 0

'le chemin du fichier alislap.ini est dans utilrepa.txt : hopale
'lecture utilrepa.txt : adresse alislap.ini
   varalis = ""
   On Error Resume Next
   Close #1: Open varutil For Input As #1
   If Err = 0 Then
      Do While Not EOF(1)
         'jld20040816
         On Error Resume Next
      
         Line Input #1, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE lecalislap 1"
            Err = 0
            Exit Do
         End If
         
         If Trim$(UCase$(a$)) = "[ALISLAP]" Then
            varint = 0
            varalis = ""
            Do While Not EOF(1)
               'jld20040816
               On Error Resume Next
            
               Line Input #1, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE lecalislap 2"
                  Err = 0
                  Exit Do
               End If

               If Trim$(a$) = "" Or Left$(Trim$(a$), 1) = "[" Then
                  Exit Do
               End If
               varstr = Trim$(UCase$(f_champ(a$, "=", 1)))
               If varstr = "PATH" And varalis = "" Then
                  varalis = Trim$(UCase$(f_champ(a$, "=", 2)))
               End If
               
               'jld20011107 les services sont définis dans [WIN_LAP]
               'If Trim$(UCase$(f_champ(b$, "=", 1))) = "SERVICES" Then
               '   varser = Trim$(f_champ(b$, "=", 2))
               'End If

               'jld20031207 vbreseau
               If varstr = "VBRESEAU" Then
                  vbreseau = Trim$(f_champ(a$, "=", 2))
                  'jld20040105
                  If Trim$(UCase$(Left$(vbreseau, 4))) = "PWD\" Then
                    vbreseau = CurDir$ & Mid$(vbreseau, 4)
                  End If
               End If

            Loop
         End If
      Loop
   Else
      Err = 0
   End If
   Close #1
   On Error GoTo 0

    If Trim$(varalis) <> "" Then
      varfic = varalis
    Else


         varwindir = Environ("windir")
         
         If varwindir = "" Then
            'jld20021024
            On Error Resume Next
             varstr = "" 'jld
             varstr = Dir("c:\windows", 16)
             If varstr = "" Then
                  'jld20021024
                  On Error Resume Next
                 varstr = Dir("c:\winnt", 16)
             End If
             If varstr = "" Then
               MsgBox "impossible de trouver le répertoire de Windows!"
               End
             Else
               varwindir = varstr
             End If
         
         End If
   
         varfic = varwindir & "\alislap.ini"

      End If

   'jld20021217 élimination des "
   varfic = f_remplace(varfic, """", "", "T", 1)


   On Error Resume Next
   Close #1: Open varfic For Input As #1
   If Err Then
        On Error Resume Next
        Close #1: Open varfic For Output As #1
        If Err Then
           MsgBox Str(Err) & " erreur ecriture : " & f_basname(varfic)
           Close #1
           Err = 0
           End
        Else
            'jld20021024
            On Error Resume Next
            varrep = "" 'jld
            varrep = Dir("c:\caralaps", 16)
            If varrep = "" Then
               ficreso = "C:\VB\LAPRESO"
               Print #1, "[WIN_LAP]"
               Print #1, "PATH=C:\VB\LAPRESO,WIN_LAP"
            Else
               ficreso = "C:\CARALAPS\VB\LAPRESO"
               Print #1, "[WIN_LAP]"
               Print #1, "PATH=C:\CARALAPS\VB\LAPRESO,CARALAPS"
            End If
            Close #1
        End If
        Err = 0
   Else
       find = False
       ficreso = ""
       varnbrbas = 0
       varnbrser = 0
       Do While Not EOF(1)
            'jld20040816
            On Error Resume Next
       
           Line Input #1, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE lecalislap 3"
                  Err = 0
                  Exit Do
               End If
               
               If Trim$(UCase$(a$)) = "[WIN_LAP]" Then
                   find = True
                   varnbr = 0 'jld20011107
                   'jld20020722 raz varnbrbas
                   'varnbrbas = 0
                   'varnbrser = 0
                   Do While Not EOF(1)
                        'jld20040816
                        On Error Resume Next
                   
                       Line Input #1, b$
                        If Err Then 'jld20011203
                           MsgBox Str(Err) & " ERREUR WHILE lecalislap 4"
                           Err = 0
                           Exit Do
                        End If

                           If Left$(Trim(b$), 1) <> "[" Then
                               If InStr(1, b$, "PATH=", 1) <> 0 And ficreso = "" Then
                                   varrep = Trim$(f_champ(b$, "=", 2))
                                    ficreso = Trim$(f_champ(varrep, varsep, 1))
                               End If
                               
                               If Trim$(UCase$(f_champ(b$, "=", 1))) = "SERVICES" Then
                                    k = 0
                                    varrep = Trim$(f_champ(b$, "=", 2))
                                    varnbrser = 0
                                    varnbrser = f_remplit_str(varrep, varsep)
                                    If varnbrser > 0 Then
                                       For i = 1 To varnbrser
                                          'jld20020715 doublon de service
                                          'ReDim Preserve tabser(i)
                                          'tabser(i) = Trim$(UCase$(tabstr(i)))
                                          'jld20020715 vérification doublon
                                          find2 = False
                                          For j = 1 To UBound(tabser)
                                             If Trim$(UCase$(f_champ(tabser(j), " ", 1))) = Trim$(UCase$(f_champ(tabstr(i), " ", 1))) Then
                                                find2 = True
                                                Exit For
                                             End If
                                          Next j
                                          If find2 = False Then
                                             k = k + 1
                                             ReDim Preserve tabser(k)
                                             tabser(k) = Trim$(UCase$(tabstr(i)))
                                          End If
                                       Next i
                                    End If
                               End If

                              'jld20011107 gestion doc
                               If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCSER" Then
                                    varrep = Trim$(f_champ(b$, "=", 2))
                                    varnbr = 0
                                    varnbr = f_remplit_str(varrep, varsep)
                                    If varnbr > 0 Then
                                       For i = 1 To varnbr
                                          ReDim Preserve tabdocser(i)
                                          tabdocser(i) = Trim$(UCase$(tabstr(i)))
                                       Next i
                                    End If
                               End If

                               If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCPRA" Then
                                    varrep = Trim$(f_champ(b$, "=", 2))
                                    varnbr = 0
                                    varnbr = f_remplit_str(varrep, varsep)
                                    If varnbr > 0 Then
                                       For i = 1 To varnbr
                                          ReDim Preserve tabdocpra(i)
                                          tabdocpra(i) = Trim$(UCase$(tabstr(i)))
                                       Next i
                                    End If
                               End If

                               If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCMOD" Then
                                    varrep = Trim$(f_champ(b$, "=", 2))
                                    varnbr = 0
                                    varnbr = f_remplit_str(varrep, varsep)
                                    If varnbr > 0 Then
                                       For i = 1 To varnbr
                                          ReDim Preserve tabdocmod(i)
                                          tabdocmod(i) = Trim$(UCase$(tabstr(i)))
                                       Next i
                                    End If
                               End If

                               If Trim$(UCase$(f_champ(b$, "=", 1))) = "DOCIDE" Then
                                    varrep = Trim$(f_champ(b$, "=", 2))
                                    varnbr = 0
                                    varnbr = f_remplit_str(varrep, varsep)
                                    If varnbr > 0 Then
                                       For i = 1 To varnbr
                                          ReDim Preserve tabdocide(i)
                                          tabdocide(i) = Trim$(UCase$(tabstr(i)))
                                       Next i
                                    End If
                               End If


                           Else
                               a$ = b$
                               Exit Do
                           End If
                   Loop
               End If
       Loop
       Close #1
       On Error GoTo 0

       If find = False Then
         On Error Resume Next
         Close #1: Open varfic For Append As #1
         If Err Then
            ficreso = "C:\VB\LAPRESO"
               MsgBox Str(Err) & " erreur ecriture : " & f_basname(varfic)
               Close #1
               Err = 0
               End
         Else
            'jld20021024
            On Error Resume Next
            varrep = "" 'jld
            varrep = Dir("c:\caralaps", 16)
            If varrep = "" Then
               ficreso = "C:\VB\LAPRESO"
               Print #1, "[WIN_LAP]"
               Print #1, "PATH=C:\VB\LAPRESO,WIN_LAP"
            Else
               ficreso = "C:\CARALAPS\VB\LAPRESO"
               Print #1, "[WIN_LAP]"
               Print #1, "PATH=C:\CARALAPS\VB\LAPRESO,CARALAPS"
            End If
            Close #1
         End If
         Close #1
         On Error GoTo 0

       End If 'find
   End If
   Close #1
   On Error GoTo 0
                            
   f_lecalislap = ficreso

End Function

Function f_lecessai2() As String
'descri= réinitialisation des variables dossier
'retour = OK

   Dim sss0 As Long
   Dim sss1 As Long
   
   Dim varstr As String
   Dim varfic As String
   Dim nf As Integer

   f_lecessai2 = ""

      'jld20021118 aider
      If lapmmi = 2 Then
         varfic = laproot & "numerow"
      Else
         'jld20040413 masqnorm multienvironnement
         'varfic = laproot & "essai2"
         varfic = laproot & lapmsq & "essai2"
      End If
         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf 'jld20031203 : Lock #nf
      If Err Then
         MsgBox Str$(Err) & " ERREUR LEC fichier : " & f_basname(varfic)
         'on garde lancien lapdos = ""
            'jld20031203 Unlock #nf
            Close #nf
            Err = 0
            Exit Function
      Else
         Line Input #nf, varstr
         lapdos = Trim$(f_champ(varstr, varsep, 1))
      End If
      'jld20031203 Unlock #nf
      Close #nf
      On Error GoTo 0

      'jld20011119
      varstr = ""
      varstr = f_lecnum(lapdos)
      If varstr <> "OK" Then
         'jld20020425
         'MsgBox "ERREUR LECTURE dossier patient : " & lapdos
         Exit Function
      End If

      f_lecessai2 = "OK"

End Function

Function f_lecnum(varstr1 As String) As String
'jld20011119
'descri= lecture dossier patient : réinitialisation des variables dossier
'varstr1 = numéro patient
'retour = OK

   Dim sss0 As Long
   Dim sss1 As Long
   
   Dim varnum As String
   Dim varstr As String
   Dim varfic As String
   Dim nf As Integer

   f_lecnum = ""

   varnum = Trim$(varstr1)
   If varnum = "" Then Exit Function

      'FRED20021118, gestion fusion
      varnum = f_fusion(varnum)
      If Trim$(varnum) = "" Then Exit Function

      lapage = ""
      lapnom = ""
      lappre = ""
      lapnai = ""
      lapsex = ""
      lapnna = ""

      'jld20020805 millier
      Select Case lapnmi
      Case 1
         lappat = f_1000(varnum, lapnum)
      Case Else
         lappat = varnum
      End Select

      'varfic = lapnum & varnum
      varfic = lapnum & lappat

      nf = FreeFile
      On Error Resume Next
      Close #nf: Open varfic For Input As #nf 'jld20031203 : Lock #nf
      If Err Then
         'jld20030331
         'jld20040209 ajout erreur 75 : fichier inexistant avec vb6
         If Err <> 53 And Err <> 52 And Err <> 75 Then
            MsgBox Str(Err) & " ERREUR LECTURE dossier : " & f_basname(varfic)
         End If
         'jld20031203 Unlock #nf
         Close #nf
         Err = 0
         Exit Function
      Else
         Line Input #nf, lapnom
         Line Input #nf, lappre
         Line Input #nf, lapnai
         Line Input #nf, lapsex
         On Error Resume Next
         Line Input #nf, varstr
         On Error Resume Next
         Line Input #nf, lapnna
      End If
      'jld20031203 Unlock #nf
      Close #nf
      On Error GoTo 0

      Err = 0
      lapdos = varnum
      lapnom = ascii_ansi(Trim$(UCase$(lapnom)))
      lappre = ascii_ansi(Trim$(UCase$(lappre)))
      lapnai = ascii_ansi(Trim$(f_temps(lapnai, 2)))
      lapsex = ascii_ansi(Trim$(UCase$(lapsex)))
      varstr = ascii_ansi(Trim$(varstr))
      lapnna = ascii_ansi(Trim$(UCase$(lapnna)))
      
      If Trim$(varstr) <> "" And Trim$(f_champ(varstr, " ", 1)) <> lapdos Then
         MsgBox "ERREUR NUMERO DOSSIER dans le fichier NUM : " & lapdos & "  " & varstr
      End If

      'FRED20030321
      'lapage = Trim$(Str$((Val(f_temps("01" + Left$(Date$, 2) + Right$(Date$, 4), 15)) - Val(f_temps(lapnai, 15))) \ 365))
      If globnote <> "" And lapapp = "AGENDACB" Then
         lapage = Trim$(Str$(Int((Val(f_temps(f_temps(globnote, 2), 15)) - Val(f_temps(lapnai, 15))) / 365.2)))
      Else
         lapage = Trim$(Str$(Int((Val(f_temps(Mid$(Date$, 4, 2) + Left$(Date$, 2) + Right$(Date$, 4), 15)) - Val(f_temps(lapnai, 15))) / 365.2)))
      End If
      If Val(lapage) < 0 Or Val(lapage) > 150 Then
         lapage = ""
      End If

      
      
      f_lecnum = "OK"
      
End Function

Function f_lecpar1000(varstr1 As String) As String
'jld20020805 lecture paramètres millier 1000
'varstr1 =
'retour = OK

   Dim varlec As Integer
   Dim varstr As String
   Dim varrep As String
   Dim nf As Integer
   Dim a$

   f_lecpar1000 = ""

   'jld20020805 millier
   varlec = 1
   lapmmi = 0
   lapnmi = 0
   
   On Error Resume Next
   nf = FreeFile
   'jld20021118 voir si adresse alisutil.txt dans lapvb ou pas
   If Trim$(lapvb) <> "" Then
      Close #nf: Open lapvb & "alisutil.txt" For Input As #nf
   Else
      Close #nf: Open "alisutil.txt" For Input As #nf
   End If
   If Err = 0 Then
      Do While Not EOF(nf)
         If varlec = 1 Then
            Line Input #nf, a$
         Else
            varlec = 1
         End If
         If Left$(Trim$(UCase$(a$)), 1) = "[" And InStr(UCase$(a$), "MILLIER") > 0 Then
            Do While Not EOF(nf)
               Line Input #nf, a$
               a$ = f_champ(a$, varsep, 1)
               If Left$(Trim$(UCase$(a$)), 1) = "[" Then
                  varlec = 0
                  Exit Do
               End If
               varstr = ""
               varrep = ""
               varstr = Trim$(UCase$(f_champ(a$, "=", 1)))
               varrep = Trim$(UCase$(f_champ(a$, "=", 2)))
               Select Case varstr
               Case "MED"
                  If varrep = "O" Or varrep = "1" Then
                     lapmmi = 1
                  End If
                  If varrep = "A" Or varrep = "2" Then
                     lapmmi = 2
                  End If
               Case "NUM"
                  If varrep = "O" Or varrep = "1" Then
                     lapnmi = 1
                  End If
               Case Else
               End Select
            Loop
         End If
      Loop
   End If
   Close #nf

   'jld20020805 millier
   'test presence dossier dans med et num
   'If lapmmi = 1 Then
   '   varstr = Dir(lapmed & "*.std", 0)
   '   If varstr <> "" Then
   '      MsgBox "ATTENTION paramètre Millier et présence dossiers dans MED"
   '      Reset
   '      End
   '   End If
   'End If

   'If lapnmi = 1 Then
   '   varstr = Dir(lapnum & "*.", 0)
   '   If varstr <> "" Then
   '      MsgBox "ATTENTION paramètre Millier et présence dossiers dans NUM"
   '      Reset
   '      End
   '   End If
   'End If

   'jld20040621
   mapmmi = lapmmi
   varmmi = lapmmi
   
   f_lecpar1000 = "OK"

End Function

Function f_lecsec(varstr1 As String, varstr2 As String) As Integer
'Sécurité : lecture fichier secure ou secure.agd pour attribution niveau
'varstr1 = login
'vartsr2 = appli

    Dim varstr As String
    Dim varapp As String
    Dim varlog As String
    Dim varpas As String
    Dim varper As Integer
    Dim varasc As Integer
    Dim identper As String
    Dim ident As String
    Dim niv As String
    Dim bool As Integer
    Dim nf As Integer
    Dim a$

    f_lecsec = 0

    varlog = Trim$(varstr1)
    varapp = Trim$(UCase$(varstr2))
    varper = 0

   If Trim$(varlog) = "" Then
      Exit Function
   End If

      bool = False

      If Left$(Trim$(UCase$(varapp)), 6) = "AGENDA" Then

       nf = FreeFile
       On Error Resume Next
       Close #nf: Open lapdic + "secure.agd" For Input As #nf 'jld20031203 : Lock #nf
         If Err Then
            Close #nf
            Screen.MousePointer = 0
            'affiche_msg_err Err, lapdic + "secure."
            Err = 0
         Else
         
             'jld20020628 si secure.agd existe on ne lit pas secure
             bool = True
            
            'On recherche le niveau d'accès en fonction de
            'l identifiant du service
            Do While Not EOF(nf)
               'jld20040816
               On Error Resume Next
            
               Line Input #nf, a$
               If Err Then
                  MsgBox Str(Err) & " ERREUR WHILE login"
                  Err = 0
                  Exit Do
               End If
               
               varasc = Asc(Left$(Trim$(a$), 1))
               Err = 0 'jld20040816 err 5 si a$ vide : asc
                  
                  If (varasc And 128) = 128 Then
                      a$ = f_cry2(a$)
                  End If
               
               ident = Trim$(f_champ(a$, varsep, 1))
               niv = Trim$(f_champ(a$, varsep, 2))

               'jld20020628
               varstr = ""
               varstr = Trim$(UCase$(f_champ(a$, varsep, 3)))
               If varstr = "" Then varstr = Trim$(UCase$(f_champ(varapp, ":", 2)))

               If Trim$(ident) = varlog And varstr = Trim$(UCase$(f_champ(varapp, ":", 2))) Then
                     bool = True
                     varper = Val(niv)
                     Exit Do
               End If
            Loop
         End If
       'jld20031203 Unlock #nf
       Close #nf
       On Error GoTo 0

      End If

      If bool = False Then

       nf = FreeFile
       On Error Resume Next
       Close #nf: Open lapdic + "secure." For Input As #nf 'jld20031203 : Lock #nf
         If Err Then
            Close #nf
            Screen.MousePointer = 0
            'affiche_msg_err Err, lapdic + "secure."
            Err = 0
            Exit Function
         End If
         Do While Not EOF(nf)
            'jld20040816
            On Error Resume Next
         
            Line Input #nf, a$
            If Err Then
               MsgBox Str(Err) & " ERREUR WHILE login"
               Err = 0
               Exit Do
            End If
            
            varasc = Asc(Left$(Trim$(a$), 1))
            Err = 0 'jld20040816 err 5 si a$ vide : asc
               
               If (varasc And 128) = 128 Then
                   a$ = f_cry2(a$)
               End If
            
            ident = Trim$(f_champ(a$, varsep, 1))
            niv = Trim$(f_champ(a$, varsep, 2))

            If Trim$(ident) = varlog Then
                  bool = True
                  varper = Val(niv)
                  Exit Do
            End If
         Loop
       'jld20031203 Unlock #nf
       Close #nf
       On Error GoTo 0

      End If

   'jld20030517 couleur verte si administrateur
   'If lapapp = "AGENDACB" Then
   '   On Error Resume Next
   '   If varper < 9 Then
   '      semaine.Label3.BackColor = QBColor(15)
   '   Else
   '      semaine.Label3.BackColor = QBColor(10)
   '   End If
   'End If

   f_lecsec = varper

End Function

Function f_ligne(varstr1 As String, varstr2 As String, varint1 As Integer) As String
'jld20031017
'permet de recomposer une ligne à partir de tabstr = ligne découpée par f_remplit_str
'varstr1 = type de tableau : tabstr,tabint,tablon,tabvar
'varstr2 = séparateur de champ
'varint1 = nombre de champ à prendre en compte

   Dim varstr As String
   Dim varlig As String
   Dim vartab As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim varchp As String
   Dim varsepfic As String
   Dim varmax As Integer
   Dim i As Integer

   f_ligne = ""

   vartab = Trim$(UCase$(varstr1))
   varsepfic = varstr2
   
   varmax = 0
   varstr = ""
   varlig = ""
   varnbr = varint1
   
   Select Case vartab
   Case "TABSTR"
      If varnbr > UBound(tabstr) Then
         varnbr = UBound(tabstr)
      End If
      
      For i = 1 To varnbr
         If i = 1 Then
            varlig = tabstr(i)
         Else
            varlig = varlig & varsepfic & tabstr(i)
         End If
      Next i
 
   Case "TABINT"
   Case "TABLON"
   Case "TABVAR"
   Case Else
   End Select
   
   f_ligne = varlig
  
End Function

Function f_log(varfic1 As String, varins1 As String, varini1 As Integer) As String
   
'Gestion des codes erreurs
'  varfic = Adresse du fichier log
'  varins = Message a inscrire
'  varini = Initialisation, 0 pour Append, 1 pour Output

   Dim varfic As String
   Dim varins As String
   Dim varini As Integer
   Dim varerr As Integer
   Dim nf As Integer

   varfic = varfic1
   varins = varins1
   varini = varini1
   varerr = Err
   
   On Error Resume Next

   f_log = "OK"
   nf = FreeFile

   Select Case varini
      Case 0
         Close #nf: Open varfic For Append As #nf
      Case 1
         Close #nf: Open varfic For Output As #nf
      Case Else
         Close #nf: Open varfic For Append As #nf
   End Select

   If Err = 0 Then
      Print #nf, Format$(Now, "DDMMYYYY") + " " + Time$ + " : Err = " + CStr(varerr) + " : " + varins
   Else
      f_log = "PB"
   End If
   Close #nf

End Function

Function f_login(varstr1 As String, varstr2 As String) As Integer
'varstr1 = login
'vartsr2 = password

    Dim varasc As Integer
    Dim identper As String
    Dim ident As String
    Dim niv As String
    Dim bool As Integer
    Dim nf As Integer
    Dim a$

    f_login = 0

    laplog = Trim$(varstr1)
    lappas = Trim$(varstr2)
    lapper = 0

   'jld20010727 on ne passe pas même si ligne blanche dans iid.service
   If Trim$(laplog) = "" Or Trim$(lappas) = "" Then
      Exit Function
   End If

    nf = FreeFile
    On Error Resume Next
    Close #nf: Open lapdic + "secure." For Input As #nf 'jld20031203 : Lock #nf
      If Err Then
         'affiche_msg_err Err, lapdic + "secure."
         Screen.MousePointer = 0
         Close #nf
         Err = 0
         Exit Function
      End If
      
      'On recherche le niveau d'accès en fonction de
      'l identifiant du service
      bool = False
      Do While Not EOF(nf)
         'Input #nf, ident, niv
         'jld20010925
         'jld20040816
         On Error Resume Next
         
         Line Input #nf, a$
         If Err Then 'jld20011203
            MsgBox Str(Err) & " ERREUR WHILE login"
            Err = 0
            Exit Do
         End If
         
         varasc = Asc(Left$(Trim$(a$), 1))
         Err = 0 'jld20040816 err 5 si a$ vide : asc
         
            If (varasc And 128) = 128 Then
                a$ = f_cry2(a$)
            End If
         
         ident = Trim$(f_champ(a$, varsep, 1))
         niv = Trim$(f_champ(a$, varsep, 2))

         If Trim$(ident) = laplog Then
               bool = True
               lapper = Val(niv)
               Exit Do
         End If
      Loop
    'jld20031203 Unlock #nf
    Close #nf
    On Error GoTo 0

   'L'identifiant du service est juste
    If bool Then
        nf = FreeFile
        On Error Resume Next
        Close #nf: Open lapdic + "iid." + lapser For Input As #nf 'jld20031203 : Lock #nf
         If Err Then
               'affiche_msg_err Err, lapdic + "iid." + lapser
               Screen.MousePointer = 0
               Close #nf
               Err = 0
               Exit Function
         End If
         'On vérifie l'identifiant personnel
         bool = False
         'Do While Not EOF(nf) And Not bool
         Do While Not EOF(nf)
               'Input #nf, identper
               'jld20010925
               'jld20040816
               On Error Resume Next
               
               Line Input #nf, a$
               If Err Then 'jld20011203
                  MsgBox Str(Err) & " ERREUR WHILE login 2"
                  Err = 0
                  Exit Do
               End If
               
               varasc = Asc(Left$(Trim$(a$), 1))
               Err = 0 'jld20040816 err 5 si a$ vide : asc
               
                  If (varasc And 128) = 128 Then
                      a$ = f_cry2(a$)
                  End If
               
               identper = Trim$(f_champ(a$, varsep, 1))
               If Trim$(identper) = lappas Then
                  bool = True
                     'jld20021024
                  If lapper = 0 Then lapper = 1
                  Exit Do
               End If
         Loop
        'jld20031203 Unlock #nf
        Close #nf
        On Error GoTo 0
    End If
    If bool = False Then
      Exit Function
    End If

   f_login = lapper

End Function

Function f_madr(varstr1 As String, varstr2 As String, varstr3 As String) As String
'traitement des adresses par mois ( accès rapide au répertoire )
'varstr1 = date 8 ex : 31122002 = 2002\200212
'varstr2 = Adresse chemin avec AntéSlach ex : lapspe ou c:\spe\
'varstr3 = fichier ex : "fichier.txt"
'retour  = 2002\200212\fichier.txt

   Dim varvar As Variant
   Dim vardat As String
   Dim varfic As String
   Dim vardir As String
   Dim varadr As String
   Dim varstr As String
   Dim varrep As String
   Dim varlon As Long
   Dim varint As Integer
   Dim varfin As String
   
   f_madr = ""

   vardat = Trim$(varstr1)
   vardir = Trim$(varstr2)
   varfic = Trim$(varstr3)

   If vardat = "" Then Exit Function
   If vardir = "" Then Exit Function
   'If varfic = "" Then Exit Function

   varvar = vardat
   If Not IsNumeric(varvar) Then
      Exit Function
   End If

   vardat = f_temps(vardat, 3)
   If vardat = "" Then
      Exit Function
   End If
   vardat = Left$(vardat, 6)

   If Right$(Trim$(vardir), 1) <> "\" Then
         vardir = vardir & "\"
   End If
   
   varadr = Left$(vardat, 4) & "\" & vardat & "\"
   varrep = f_mkdir(vardir & varadr, "", 0)
   If varrep <> "OK" Then
      Exit Function
   End If

   f_madr = varadr & varfic

End Function

Function f_maj_essai2(varstr1 As String, varstr2 As String, varstr3 As String) As String
'varstr1 = adresse essai2 : c:\
'varstr2 = numéro de dossier
'varstr3 = R = Lecture ou W = Ecriture

   Dim varact As String
   Dim varstr As String
   Dim varrep As String
   Dim varfic As String
   Dim varlig As String
   Dim varnumdos As String
   Dim nf As Integer

   f_maj_essai2 = ""
  
   varact = UCase$(Trim$(varstr3))
   If varact <> "R" And varact <> "W" Then
      MsgBox ("ERREUR PARAMETRES maj essai2 3")
      Exit Function
   End If
   

   varnumdos = Trim$(varstr2)
   'jld20020502
   'If varnumdos = "" And varact = "W" Then
   'jld20040421 attention dossier PDJ
   'If Val(varnumdos) = 0 And varact = "W" Then
   If (Trim$(varnumdos) = "" Or varnumdos = "0") And varact = "W" Then
      MsgBox ("ERREUR PARAMETRES maj essai2 2")
      Exit Function
   End If

   varrep = Trim$(varstr1)
   If varrep = "" Then
      MsgBox ("ERREUR PARAMETRES maj essai2 1")
      Exit Function
   End If

      'jld20021118 aider
      If lapmmi = 2 Then
         varfic = varrep & "numerow"
      Else
         varfic = varrep & "essai2"
      End If
         
   If varact = "R" Then
      Err = 0
      nf = FreeFile
      On Error Resume Next
      Close #nf: Open varfic For Input As #nf 'jld20031203 : Lock #nf
      If Err Then
         MsgBox (Str(Err) & " ERREUR LEC fichier : " & f_basname(varfic))
         Close #nf
         Err = 0
         Exit Function
      End If
      Line Input #nf, varlig
      'jld20031203 Unlock #nf
      Close #nf
      On Error GoTo 0

      varnumdos = Trim$(varlig)
   End If

   If varact = "W" Then
      Err = 0
      nf = FreeFile
      On Error Resume Next
      Close #nf: Open varfic For Output Lock Read Write As #nf  'jld20031203 : Lock #nf
      If Err Then
         MsgBox (Str(Err) & " ERREUR ECR fichier : " & f_basname(varfic))
         'jld20031224 Unlock #nf
         Close #nf
         Err = 0
         Exit Function
      End If
      Print #nf, varnumdos
      'jld20031224 Unlock #nf
      Close #nf
      On Error GoTo 0

      'jld20021118 aider
      If lapmmi = 2 Then
         On Error Resume Next
         FileCopy varfic, "c:\numerow"
      Else
         On Error Resume Next
         FileCopy varfic, "c:\essai2"
      End If

      On Error GoTo 0

      Err = 0
   End If

      f_maj_essai2 = varnumdos

End Function

Function f_mkdir(varstr1 As String, varstr2 As String, varint1 As Integer) As String
'jld20020122 test presence repertoire et création des répertoires manquants
'varstr1 = adresse repertoire sans le fichier
'varstr2 = futur
'varint1 = 1 : avec message de validation
'retour = 9999 annulation création
'retour = 9998 adresse non valide
'retour = err   OK

   Dim varerr As Integer
   Dim varadr As String
   Dim varfic As String
   Dim varrep As String
   Dim varstr As String
   Dim varnom As String
   Dim vardir As String
   Dim varcur As String
   Dim varmes As Integer

   f_mkdir = ""

   varadr = varstr1
   varmes = varint1
   varerr = 0
   
   vardir = ""
   varfic = ""

   On Error Resume Next
   varrep = ""
      
   varcur = CurDir
   
   Do While True
      
         
      'futur à développer ?
      'pour fichier vardir = Trim$(f_dirname(varadr))
      'pour fichier varnom = Trim$(f_basname(varadr))
      
      vardir = Trim$(varadr)

      If vardir = "" Then
         varerr = 9998
         Exit Do
      End If

      If Right$(vardir, 1) = "\" Then
         vardir = Left$(vardir, Len(vardir) - 1)
      End If

      'jld20020805 attention aux fichier avec le même nom que le repertoire
      'jld20021024
      On Error Resume Next
      varrep = ""
      varrep = Dir(vardir, 16)
      If varrep <> "" Then
         If (GetAttr(vardir) And 16) <> 16 Then
            varerr = 9997
            MsgBox "ATTENTION CREATION REPERTOIRE IMPOSSIBLE : fichier avec même nom que repertoire : " & vardir
         End If
         Exit Do
      End If

      Do While True
         varstr = vardir
         
         vardir = Trim$(f_dirname(varstr))
         varnom = Trim$(f_basname(varstr))
         
         If vardir = "" Then
            varerr = 9998
            Exit Do
         End If
         If Right$(vardir, 1) = "\" Then
            vardir = Left$(vardir, Len(vardir) - 1)
         End If
         'jld20021024
         On Error Resume Next
         varrep = ""
         varrep = Dir(vardir, 16)
         If varrep <> "" Then
            If varmes = 1 Then
               varrep = ""
               varrep = InputBox("REPERTOIRE INEXISTANT : " & vardir & "\" & varnom & Chr$(13) & Chr$(10) & "Voulez vous le créer (O/N) : ", "Création répertoire", "O")
               If Trim$(UCase$(varrep)) <> "O" Then
                  varerr = 9999
                  Exit Do
               End If
            End If
            On Error Resume Next
            MkDir (vardir & "\" & varnom)
            If Err Then
               varerr = Err
            End If
            Exit Do
         End If
      Loop
      If varerr <> 0 Then Exit Do
   Loop

   If varerr = 0 Then
      f_mkdir = "OK"
   Else
      f_mkdir = Str$(varerr)
   End If

End Function

Function f_qui_fait_quoi(fictrace As String, varmsg As String) As String
' SB20020127  TRACE LE VB/LOGIN.TXT SI EXISTE+UN MESSAGE AU CHOIX
   
   Dim varstr As String
   Dim varint As Integer
   Dim i As Integer

   Dim ficlogin As String
   
   Dim msg As String
   Dim trace As String
   
   Dim canalin As Integer
   Dim canalout As Integer
   
   canalin = FreeFile
   ficlogin = lapvb & "login.txt"
   
   'On lit le login
   
   On Error Resume Next
   
   msg = ""
   Open ficlogin For Input As #canalin
      If Err Then
         Err = 0
      Else
         Line Input #canalin, msg
         'jld20020418
         'date login,heure,profile,metier,grade
         '13062001,12:21,TEST,DIET,8
         '    8   ,  5  , 8  ,  8 ,1
         varint = f_remplit_str(msg, varsep)
         If varint > 5 Then
            varint = 5
         End If
         For i = 1 To varint
            Select Case i
            Case 1
               varstr = f_r_d(f_temps(tabstr(i), 2), " ", 8)
            Case 2
               varstr = varstr & varsep & f_r_d(tabstr(i), " ", 5)
            Case 3
               varstr = varstr & varsep & f_r_d(tabstr(i), " ", 8)
            Case 4
               varstr = varstr & varsep & f_r_d(tabstr(i), " ", 8)
            Case 5
               varstr = varstr & varsep & f_r_d(tabstr(i), " ", 1)
            Case Else
            End Select
         Next i
         
         For i = varint + 1 To 5
            Select Case i
            Case 1
               varstr = f_r_d(" ", " ", 8)
            Case 2
               varstr = varstr & varsep & f_r_d(" ", " ", 5)
            Case 3
               varstr = varstr & varsep & f_r_d(" ", " ", 8)
            Case 4
               varstr = varstr & varsep & f_r_d(" ", " ", 8)
            Case 5
               varstr = varstr & varsep & f_r_d(" ", " ", 1)
            Case Else
            End Select
         Next i

         msg = varstr

         msg = msg & ","

      End If
      trace = msg & varmsg
   
   Close #canalin
    
    'On écrit la trace
   canalout = FreeFile
   Open fictrace For Append As #canalout
      If Err Then
         Err = 0
         Exit Function
      End If
   
      
      Print #canalout, trace
   
   Close #canalout
   
   On Error GoTo 0

End Function

Function f_r_d(varstr1 As String, varstr2 As String, varvar1 As Variant) As String
'DESCRI=remplissage de données à droite : f_r_d(varlig,varcar,nbrcar:0/1)
'varstr1 = donnee a completer
'varstr2 = caractere de complement
'varvar1 = longueur a obtenir  : 1=on ne coupe pas

   Dim varlig As String
   Dim varstr As String
   Dim varcar As String
   Dim varlen As Integer
   Dim varcou As Integer

   varlig = varstr1
   varcar = varstr2
   varstr = varvar1

   If InStr(varstr, ":") = 0 Then
      varlen = Val(varstr)
      varcou = 0
   Else
      varlen = Val(Left$(varstr, InStr(varstr, ":") - 1))
      varcou = Val(Mid$(varstr, InStr(varstr, ":") + 1))
   End If

   f_r_d = varlig

   If varcou <> 1 Then
      varlig = Left$(varlig, varlen)
   End If

   'jld20030710
   If varcar <> "" Then
      Do While Len(varlig) < varlen
          varlig = varlig & varcar
      Loop
   End If

   f_r_d = varlig

End Function

Function f_r_g(varstr1 As String, varstr2 As String, varvar1 As Variant) As String
'DESCRI=remplissage de données à gauche : f_r_g(varlig,varcar,nbrcar:0/1)
'varstr1 = donnee a completer
'varstr2 = caractere de complement
'varvar1 = longueur a obtenir : x (si x = 1 : on ne coupe pas)

   Dim varlig As String
   Dim varstr As String
   Dim varcar As String
   Dim varlen As Integer
   Dim varcou As Integer

   varlig = varstr1
   varcar = varstr2
   
   varstr = varvar1

   If InStr(varstr, ":") = 0 Then
      varlen = Val(varstr)
      varcou = 0
   Else
      varlen = Val(Left$(varstr, InStr(varstr, ":") - 1))
      varcou = Val(Mid$(varstr, InStr(varstr, ":") + 1))
   End If

   f_r_g = varlig

   If varcou <> 1 Then
      varlig = Right$(varlig, varlen)
   End If

   'jld20030710
   If varcar <> "" Then
      Do While Len(varlig) < varlen
          varlig = varcar & varlig
      Loop
   End If

   f_r_g = varlig

End Function

Function f_r_lon(varstr1 As String, varstr2 As String) As Integer
'remplit le tableau tablon
'varstr1 ligne à découper
'varstr2 séparateur
'retour = nombre de données découpées
'retour = colstr : nombre de colonne du tableau tabstr

   Dim a$
   Dim varlig As String
   Dim ccpt As Long
   ReDim tablon(0)

   varlig = varstr1
   ccpt = 0
   f_r_lon = ccpt
   colstr = ccpt

plusreglelon:

   If InStr(varlig, varstr2) <> 0 Then
        ccpt = ccpt + 1
        ReDim Preserve tablon(ccpt)
        tablon(ccpt) = Val(Left$(varlig, InStr(varlig, varstr2) - 1))
        varlig = Right$(varlig, Len(varlig) - InStr(varlig, varstr2) - Len(varstr2) + 1)
   Else
        GoTo finreglelon
   End If
        
   GoTo plusreglelon

finreglelon:

   ccpt = ccpt + 1
   ReDim Preserve tablon(ccpt)
   tablon(ccpt) = Val(varlig)
   f_r_lon = ccpt
   colstr = ccpt

End Function

Function f_r_var(varvar1 As Variant, varstr2 As String) As Integer
'remplit le tableau tabvar
'varvar1 ligne à découper
'varstr2 séparateur
'retour = nombre de données découpées
'retour = colstr : nombre de colonne du tableau tabstr

   Dim a$
   Dim varvar As Variant
   Dim varlig As String
   Dim ccpt As Long
   ReDim tabvar(0)

   varvar = varvar1
   ccpt = 0
   f_r_var = ccpt
   colstr = ccpt

plusreglevar:

   If InStr(varvar, varstr2) <> 0 Then
        ccpt = ccpt + 1
        ReDim Preserve tabvar(ccpt)
        tabvar(ccpt) = Left$(varvar, InStr(varvar, varstr2) - 1)
        varvar = Right$(varvar, Len(varvar) - InStr(varvar, varstr2) - Len(varstr2) + 1)
   Else
        GoTo finreglevar
   End If
        
   GoTo plusreglevar

finreglevar:

   ccpt = ccpt + 1
   ReDim Preserve tabvar(ccpt)
   tabvar(ccpt) = varvar
   f_r_var = ccpt
   colstr = ccpt

End Function

Function f_rechv(varstr1 As String, varstr2 As String, varstr3 As String, varint1 As Integer, varint2 As Integer, varint3 As Integer) As String
'descri = recherche verticale dans un fichier
'varstr1 = adresse fichier
'varstr2 = séparateur fichier
'varstr3 = valeur recherchée
'varint1 = colonne de recherche
'varint2 = colonne à retourner
'varint3 = test majuscules minuscule indiférent : 1
'retour = valeur trouvée

   Dim varvar As Variant
   Dim varerrcpt As Integer 'jld20020105
   Dim varfic As String
   Dim varsepfic As String
   Dim varval As String
   Dim varcol1 As Integer
   Dim varcol2 As Integer
   Dim varcas As Integer
   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim nf As Integer
   Dim a$

      
   f_rechv = ""

   
   varfic = Trim$(UCase$(varstr1))
   varsepfic = varstr2
   If Trim$(UCase$(varsepfic)) = "VIRGULE" Then varsepfic = ","
   varval = varstr3
   varcol1 = varint1
   varcol2 = varint2
   varcas = varint3
   If varcas = 1 Then
      varval = Trim$(UCase$(varval))
   End If
   
   If varfic = "" Then Exit Function
   If varsepfic = "" Then Exit Function
   If varval = "" Then Exit Function
   If varcol1 = 0 Then Exit Function
   If varcol2 = 0 Then Exit Function

   varerrcpt = 0
   varint = 0
   Do While True
   
      varerrcpt = varerrcpt + 1
      nf = FreeFile
      On Error Resume Next
      'Close #nf: Open varfic For Input Lock Read Write As #nf  'jld20031203 : Lock #nf
      Close #nf: Open varfic For Input Lock Write As #nf  'jld20041005
      
      'jld20041115
      If Err = 53 Then
         Exit Do
      End If
      If Err Then
            varint = varint + 1
            If varint > 5 Then
                  Exit Do
            Else
               varvar = Timer + 1
               While Timer < varvar
                  'DoEvents
               Wend
            End If
      Else
         Exit Do
      End If
      Close #nf

      
      'If Err Then
      '   If varerrcpt > 99 Then
      '      Close #nf
      '      Err = 0
      '      Screen.MousePointer = 0
      '      Exit Function
      '   Else
      '      Close #nf
      '      Err = 0
      '   End If
      'Else
      '   Exit Do
      'End If
   Loop
   If Err Then
      Close #nf
      Screen.MousePointer = 0
      Err = 0
      Exit Function
   End If
   
   Do While Not EOF(nf)
       'jld20040816
       On Error Resume Next
   
      Line Input #nf, a$
      If Err Then 'jld20011203
         MsgBox Str(Err) & " ERREUR WHILE rechv"
         Err = 0
         Exit Do
      End If

      varstr = ""
      varstr = f_champ(a$, varsepfic, varcol1)
      If varcas = 1 Then
         varstr = Trim$(UCase$(varstr))
      End If
      If varstr = varval Then
         varrep = f_champ(a$, varsepfic, varcol2)
         Exit Do
      End If
   Loop
   'jld20031224 Unlock #nf
   Close #nf
   On Error GoTo 0

   f_rechv = varrep

End Function

Function f_rechv2(varstr1 As String, varstr2 As String, varstr3 As String, varint1 As Integer, varstr4 As String, varint2 As Integer, varvar1 As Variant, varint4 As Integer) As String
'descri = recherche verticale dans un fichier
'varstr1 = adresse fichier  par défaut dans le dic
'varstr2 = séparateur fichier
'varstr3 = valeur recherchée cle 1
'varint1 = colonne de recherche cle 1
'varstr4 = valeur recherchée cle 2
'varint2 = colonne de recherche cle 2
'varvar1 = colonnes à retourner: 5+8
'varint4 = 0-1 : test majuscules minuscule, 1 : indiférent
'varint4 = 2 : récupère le dernier enregistrement
'varint4 = 4 : fichier crypté
'retour = valeur trouvée

'si aucune clé : lecture première ligne : exemple login.txt

   Dim varvar As Variant
   Dim vartst As Integer
   Dim varasc As Integer
   Dim varnumchp As Integer 'jld20011125
   Dim varfin As Integer
   Dim varfic As String
   Dim varsepfic As String
   Dim varval As String
   Dim varcle1 As String
   Dim varcle2 As String
   Dim varcolcle1 As Integer
   Dim varcolcle2 As Integer
   Dim varcolval As Integer
   Dim varcas As Integer
   Dim varstr As String
   Dim varrep As String
   Dim find As Integer
   Dim nf As Integer
   Dim a$
   Dim tabreccol() As Integer
   ReDim tabreccol(0)
   Dim varint As Integer
   Dim i As Integer

      
   f_rechv2 = ""

   varval = "***"
   varfic = Trim$(UCase$(varstr1))
   varsepfic = varstr2
   varcle1 = varstr3
   varcolcle1 = varint1
   varcle2 = varstr4
   varcolcle2 = varint2

   'jld20030815
   'varcolval = varint3
   varstr = varvar1
   varint = f_r_lon(varstr, "+")
   For i = 1 To varint
      ReDim Preserve tabreccol(i)
      tabreccol(i) = tablon(i)
   Next i

   'jld20010727
   If varint4 >= 4 Then 'fichier crypté
       vartst = 1
       varint4 = varint4 - 4
   Else
      vartst = 0
   End If

   'jld20010727
   If varint4 >= 2 Then 'recherche dernier
       varfin = 1
       varint4 = varint4 - 2
   Else
      varfin = 0
   End If
   
   varcas = varint4

   If varcas = 1 Then
      varcle1 = Trim$(UCase$(varcle1))
      varcle2 = Trim$(UCase$(varcle2))
   End If
   
   If varfic = "" Then Exit Function
   If varsepfic = "" Then Exit Function
   'If varcle1 = "" Then Exit Function
   'If varcolcle1 = 0 Then Exit Function

   'jld20030815 If varcolval = 0 Then Exit Function

   If InStr(varfic, "\") = 0 Then
      varfic = lapdic & varfic
   Else
      varstr = Trim$(UCase$(f_dirname(varfic)))
      If varstr = "LAPTMP\" Then
         varfic = laptmp & Trim$(UCase$(f_basname(varfic)))
      End If
      If varstr = "C:\VB\" And Trim$(UCase$(lapvb)) <> "C:\VB\" Then
         varfic = lapvb & Trim$(UCase$(f_basname(varfic)))
      End If
   End If

   If Trim$(UCase$(varsepfic)) = "VIRGULE" Then
      varsepfic = ","
   End If

   Select Case Trim$(UCase$(varcle1))
   Case "LAPAGE"
      varcle1 = lapage
   Case "LAPNOM"
      varcle1 = lapnom
   Case "LAPPRE"
      varcle1 = lappre
   Case "LAPNAI"
      varcle1 = lapnai
   Case "LAPSEX"
      varcle1 = lapsex
   Case "LAPDOS"
      varcle1 = lapdos
   Case "LAPNNA"
      varcle1 = lapnna
   End Select

   'jld20011125 recupe champ masque @
   varnumchp = 0
   If Left$(varcle1, 1) = "@" Then
      varnumchp = Val(f_champ(varcle1, "@", 2))
      If varnumchp > 0 Then
         varcle1 = donnees(9, varnumchp)
      End If
   End If

   Select Case Trim$(UCase$(varcle2))
   Case "LAPAGE"
      varcle2 = lapage
   Case "LAPNOM"
      varcle2 = lapnom
   Case "LAPPRE"
      varcle2 = lappre
   Case "LAPNAI"
      varcle2 = lapnai
   Case "LAPSEX"
      varcle2 = lapsex
   Case "LAPDOS"
      varcle2 = lapdos
   Case "LAPNNA"
      varcle2 = lapnna
   End Select

   'jld20011125 recupe champ masque
   varnumchp = 0
   If Left$(varcle2, 1) = "@" Then
      varnumchp = Val(f_champ(varcle2, "@", 2))
      If varnumchp > 0 Then
         varcle2 = donnees(9, varnumchp)
      End If
   End If


   varint = 0
   Do While True
   
   nf = FreeFile
   On Error Resume Next
   'Close #nf: Open varfic For Input Lock Read Write As #nf  'jld20031203 : Lock #nf
   Close #nf: Open varfic For Input Lock Write As #nf  'jld20041005
   
      'jld20041115
      If Err = 53 Then
         Exit Do
      End If
      If Err Then
            varint = varint + 1
            If varint > 5 Then
                  Exit Do
            Else
               varvar = Timer + 1
               While Timer < varvar
                  'DoEvents
               Wend
            End If
      Else
         Exit Do
      End If
      Close #nf
   
   Loop

   If Err Then
      'jld20031224 Unlock #nf
      Close #nf
      Err = 0
      Screen.MousePointer = 0
      Exit Function
   End If
   
   Do While Not EOF(nf)
      'jld20040816 err 5 si ligne vide
      On Error Resume Next
      Line Input #nf, a$
      If Err Then 'jld20011203
         MsgBox Str(Err) & " ERREUR WHILE rechv2"
         Err = 0
         Exit Do
      End If

      'jld20040801 cryptage
      If vartst = 1 Then
         varasc = Asc(Left$(Trim$(a$), 1))
         Err = 0 'jld20040816 err 5 si a$ vide : asc
         
         If (varasc And 128) = 128 Then
            a$ = f_cry2(a$)
         End If
         
      End If
      
      If Trim$(a$) <> "" Then
         varstr = ""
         If varcolcle1 <> 0 Then
            varstr = f_champ(a$, varsepfic, varcolcle1)
         End If
         
         'jld20040712
         'If varcas = 1 Then
         '   varstr = Trim$(UCase$(varstr))
         'End If
   
         find = False
   
         'jld20040712
         If varcas = 1 Then
            If Trim$(UCase$(varstr)) = Trim$(UCase$(varcle1)) Then
               find = True
            End If
         Else
            If varstr = varcle1 Then
               find = True
            End If
         End If
         
         varstr = ""
         If varcolcle2 <> 0 Then
            varstr = f_champ(a$, varsepfic, varcolcle2)
         End If
   
         'jld20040712
         'If varcas = 1 Then
         '   varstr = Trim$(UCase$(varstr))
         'End If
               
         'jld20040712
         If varcas = 1 Then
            If find = True And Trim$(UCase$(varstr)) = Trim$(UCase$(varcle2)) Then
               find = True
            Else
               find = False
            End If
         Else
            If find = True And varstr = varcle2 Then
               find = True
            Else
               find = False
            End If
         End If
         
         If find = True Then
            'jld20030815
            'varval = f_champ(a$, varsepfic, varcolval)
            For i = 1 To UBound(tabreccol)
               If i = 1 Then
                  varval = f_champ(a$, varsepfic, tabreccol(i))
               Else
                  varval = varval & varsep & f_champ(a$, varsepfic, tabreccol(i))
               End If
            Next i
            
            'jld20010727 recherche dernier de plusieurs si varfin = 1
            If varfin = 0 Then
               Exit Do
            End If
         End If
      End If 'a$
   
   Loop
   'jld20031224 Unlock #nf
   Close #nf
   On Error GoTo 0

   f_rechv2 = varval

End Function

Function f_remplace(varstr1 As String, varstr2 As String, varstr3 As String, varstr4 As String, varint1 As Integer) As String
'varstr1 chaine à traiter
'varstr2 chaine à remplacer
'varstr3 chaine de remplacement
'varstr4 tous
'varint1 case sensitive

   Dim varstr As String
   Dim varori As String
   Dim vardes As String
   Dim varres As String
   Dim varint As Integer
   Dim varpos As Long
   Dim varlen As Long
   Dim varLenori As Long
   Dim varcmp As Integer

   varstr = varstr1
   varori = varstr2
   vardes = varstr3

   f_remplace = varstr
   
   If varori = vardes Then
      Exit Function
   End If

   If varint1 <> 0 And varint1 <> 1 Then
      varcmp = 0
   Else
      varcmp = varint1
   End If


   varLenori = Len(varori)
   varres = ""
   'varint = len(varres)
   
   Do While InStr(1, varstr, varori, varcmp) > 0
      On Error Resume Next
      varlen = Len(varstr)
      If Err Then MsgBox Str(Err) & " ERREUR REMPLACE 1 : " & varori
      varpos = InStr(1, varstr, varori, varcmp)
      If Err Then MsgBox Str(Err) & " ERREUR REMPLACE 2 : " & varori
      varres = Left$(varstr, varpos - 1)
      If Err Then MsgBox Str(Err) & " ERREUR REMPLACE 3 : " & varori
      varres = varres & vardes
      If Err Then MsgBox Str(Err) & " ERREUR REMPLACE 4 : " & varori
      varres = varres & Right$(varstr, varlen - varpos - varLenori + 1)
      If Err Then MsgBox Str(Err) & " ERREUR REMPLACE 5 : " & varori
      varstr = varres
      varres = ""
      If UCase$(Left$(Trim$(varstr4), 1)) <> "T" Then Exit Do
   Loop
   On Error GoTo 0

   f_remplace = varstr

End Function

Function f_remplit_donnees(varint1 As Integer, varstr1 As String, varstr2 As String) As Integer
'varint1 numéro de ligne dans le tableau
'varstr1 ligne à découper
'varstr2 séparateur
'retour = nombre de données découpées

    Dim a$
    Dim varlig As String
    Dim i, ccpt As Integer

   If varint1 = 9 Or varint1 = 10 Then
      MsgBox ("ATTENTION utilisation interdite dans le tableau donnees (9-10) : " & Str(varint1))
      Reset
      End
   End If

   varlig = varstr1

   For i = 1 To UBound(donnees)
      donnees(varint1, i) = ""
   Next i

   ccpt = 0
   f_remplit_donnees = 0

pplusregle:

   If InStr(varlig, varstr2) <> 0 Then
        ccpt = ccpt + 1
        donnees(varint1, ccpt) = Left$(varlig, InStr(varlig, varstr2) - 1)
'        varlig = Right$(varlig, Len(varlig) - Len(Left$(varlig, InStr(varlig, varstr2))))
        varlig = Right$(varlig, Len(varlig) - InStr(varlig, varstr2) - Len(varstr2) + 1)
   Else
        GoTo pfinregle
   End If
        
   GoTo pplusregle

pfinregle:

   ccpt = ccpt + 1
   donnees(varint1, ccpt) = varlig
   f_remplit_donnees = ccpt

End Function

Function f_remplit_droite(varstr1 As String, varstr2 As String, varlen As Integer) As String
'varstr1 = donnee a completer
'varstr2 = caractere de complement
'varlen = longueur a obtenir

   Dim varstr As String

   varstr = varstr1

   f_remplit_droite = varstr

   varstr = Left$(Trim$(varstr), varlen)

   Do While Len(varstr) < varlen
       varstr = varstr & varstr2
   Loop

   f_remplit_droite = varstr

End Function

Function f_remplit_gauche(varstr1 As String, varstr2 As String, varlen As Integer) As String
'varstr1 = donnee a completer
'varstr2 = caractere de complement
'varlen = longueur a obtenir

   Dim varstr As String

   varstr = varstr1

   f_remplit_gauche = varstr

   varstr = Right$(Trim$(varstr), varlen)

   Do While Len(varstr) < varlen
       varstr = varstr2 & varstr
   Loop

   f_remplit_gauche = varstr

End Function

Function f_remplit_str(varstr1 As String, varstr2 As String) As Integer
'remplit le tableau tabstr
'varstr1 ligne à découper
'varstr2 séparateur
'retour = nombre de données découpées
'retour = colstr : nombre de colonne du tableau tabstr
'err

   Dim a$
   Dim varlig As String
   Dim ccpt As Long
   ReDim tabstr(0)

   varlig = varstr1
   ccpt = 0
   f_remplit_str = ccpt
   colstr = ccpt

plusregle:

   If InStr(varlig, varstr2) <> 0 Then
        ccpt = ccpt + 1
        ReDim Preserve tabstr(ccpt) 'memory full
        tabstr(ccpt) = Left$(varlig, InStr(varlig, varstr2) - 1)
        varlig = Right$(varlig, Len(varlig) - InStr(varlig, varstr2) - Len(varstr2) + 1)
   Else
        GoTo finregle
   End If
        
   GoTo plusregle

finregle:

   ccpt = ccpt + 1
   ReDim Preserve tabstr(ccpt) 'memory full
   tabstr(ccpt) = varlig
   f_remplit_str = ccpt
   colstr = ccpt

End Function

Function f_rwlap(varfic1 As String, varlec1 As String, vartit1 As String, varite1 As String, varval1 As String, vardiv1 As Variant) As String
'
'Fonction de gestion des .lap
'f_rwlap(varfic1 As String, varlec1 As String, vartit1 As String, varite1 As String, varval1 As String, vardiv1 As Variant)
'varfic1 = fichier à traiter
'varlec1 = "R" (lecture) ou "W" (ecriture)
'vartit1 = titre [] du groupe incluant l'item
'varite1 = item à modifier ou inserer
'  si varlec = "R" et varite = "" alors on lit tout le groupe, les lignes seront memorisees dans le tableau tabrwlap
'varval1 = valeur à assigner à l'item
'vardiv1 = divers, libre
'vardiv1 = 4 si le fichier est crypté
'
'f_rwlap = valeur cherchee en lecture
'          le nombre de lignes lues en recuperation de groupe : tabrwlap(varnbr)
'f_rwlap = "OK" en ecriture

   Dim vartst As Integer
   Dim varasc As Integer
   Dim varfic As String
   Dim varlig As String
   Dim varlec As String
   Dim vartit As String
   Dim varite As String
   Dim varval As String
   Dim vardiv As Variant
   
   Dim vartmp As String
   Dim varstr As String
   Dim varint As Integer
   
   Dim lect As String
   Dim lect1 As String
   Dim lect2 As String
   Dim lect3 As String
   Dim x As Integer
   
   Dim nf As Integer
   Dim nf2 As Integer
   Dim varnbr As Integer
   
   f_rwlap = "@@@"
   
   varnbr = 0
   
   'jld20040801
   If Val(vardiv1) = 4 Then
      vartst = 1
   Else
      vartst = 0
   End If
   
   On Error Resume Next

   varfic = UCase$(Trim$(varfic1))
   varlec = UCase$(Trim$(varlec1))
   vartit = UCase$(Trim$(vartit1))
   varite = UCase$(Trim$(varite1))
   
   'on efface le tableau que pour une lecture globale de famille [
   If varlec = "R" And varite = "" Then
      ReDim tabrwlap(0)
      'lect2 = "0"
      lect2 = "@@@"
   End If
   
   varval = UCase$(Trim$(varval1))
   vardiv = UCase(Trim(vardiv1))

   varfic = f_adr(varfic)
   vartmp = f_dirname(varfic) + "_" + Mid$(f_basname(varfic), 2)

   FileCopy varfic, vartmp

   If Err <> 0 Then
      'jld MsgBox "Erreur recopie du fichier"
      f_rwlap = "@@@@"
      GoTo finfic
   End If

   'Lecture
   If varlec = "R" Then
      On Error Resume Next
      nf = FreeFile
      Close #nf: Open vartmp For Input As #nf
      If Err = 0 Then
         Do While Not EOF(nf)
            Err = 0
            Line Input #nf, lect
            If Err Then
               f_rwlap = "@@@@"
               GoTo finfic
            End If
            
            'jld20040801 cryptage
            If vartst = 1 Then
               varasc = Asc(Left$(Trim$(lect), 1))
               Err = 0 'jld20040816 err 5 si a$ vide : asc
               
               If (varasc And 128) = 128 Then
                  lect = f_cry2(lect)
               End If
               
            End If

            'on enlève les commentaires
            lect = f_champ(lect, "'", 1)
            
            'on garde le texte comme il a été écrit
            'lect = UCase$(Trim$(lect))
            varlig = ""
            varlig = Trim$(UCase$(lect))
            
            If Left$(varlig, 1) <> "'" And Left$(varlig, 1) <> "_" And varlig <> "" Then
            
               'If lect = "[" + vartit + "]" Then
               If InStr(1, varlig, "[" + vartit + "]", 1) = 1 Then
                  Do While Not EOF(nf)
                     On Error Resume Next
                     Line Input #nf, lect
                     
                     If Err Then
                        f_rwlap = "@@@@"
                        GoTo finfic
                        'Exit Do
                     End If
                     
                     'jld20040801 cryptage
                     If vartst = 1 Then
                        varasc = Asc(Left$(Trim$(lect), 1))
                        Err = 0 'jld20040816 err 5 si a$ vide : asc
                        
                        If (varasc And 128) = 128 Then
                           lect = f_cry2(lect)
                        End If
                        
                     End If
                     
                     'on enlève les commentaires
                     lect = f_champ(lect, "'", 1)

                     'lect = UCase$(Trim$(lect))
                     varlig = ""
                     varlig = Trim$(UCase$(lect))
                     
                     If Left$(varlig, 1) = "[" Then
                        Exit Do
                     End If
                     
                     If Left$(varlig, 1) <> "'" And Left$(varlig, 1) <> "_" And varlig <> "" Then
                        lect1 = Trim$(UCase$(f_champ(lect, "=", 1)))
                        
                        'jld attention fred utilise le tableau tabrwlap si varite = ""
                        'varnbr = varnbr + 1
                        'ReDim Preserve tabrwlap(varnbr)
                        'tabrwlap(varnbr) = lect
                        'lect2 = CStr(varnbr)
                     
                        'If varite <> "" And lect1 = varite Then
                        '   varnbr = 1
                        '   ReDim tabrwlap(varnbr)
                        '   tabrwlap(varnbr) = lect
                        '   lect2 = f_champ(lect, "=", 2)
                        '   Exit Do
                        'End If

                        If varite <> "" And varite = lect1 Then
                           lect2 = f_champ(lect, "=", 2)
                           Exit Do
                        End If
                        If varite = "" Then
                           varnbr = varnbr + 1
                           ReDim Preserve tabrwlap(varnbr)
                           tabrwlap(varnbr) = lect
                           lect2 = CStr(varnbr)
                        End If
                        
                     End If
                  Loop
                  Exit Do
               End If
            End If
         Loop
      Else
         'MsgBox "Erreur lecture fichier"
         f_rwlap = "@@@@"
         GoTo finfic
      End If
   
      Close #nf

      f_rwlap = lect2

   End If

   'Ecriture
   If varlec = "W" Then
      On Error Resume Next
      nf = FreeFile
      Close #nf: Open varfic For Input As #nf
      nf2 = FreeFile
      Close #nf2: Open vartmp For Output As #nf2
      If Err = 0 Then
         Do While Not EOF(nf)
            Err = 0
            Line Input #nf, lect
            If Err Then
               f_rwlap = "@@@@"
               GoTo finfic
            End If
            
            lect = UCase$(Trim$(lect))
            Print #nf2, lect
          
            'If Left$(lect, 1) <> "'" And Left$(lect, 1) <> "_" And lect <> "" Then
            If lect = "[" + vartit + "]" Then
               Do While Not EOF(nf)
                  Err = 0
                  Line Input #nf, lect
                  If Err Then
                     Exit Do
                  End If
                  
                  lect = UCase$(Trim$(lect))
                  If Left$(lect, 1) = "[" Then
                     Print #nf2, lect
                     Exit Do
                  End If
         
                  'If Left$(lect, 1) <> "'" And Left$(lect, 1) <> "_" And lect <> "" Then
                     lect1 = f_champ(lect, "=", 1)
                     If lect1 = varite Then
                        lect = lect1 + "=" + varval
                        'Exit Do
                     End If
                  'End If
         
                  Print #nf2, lect
               Loop
               'Exit Do
            End If
            'End If
         Loop
      Else
         'jld MsgBox "Erreur lecture fichier"
         f_rwlap = "@@@@"
         GoTo finfic
      End If
   
      Close #nf
      Close #nf2

      f_rwlap = ""
      If Err = 0 Then
         FileCopy vartmp, varfic
         f_rwlap = "OK"
      End If

   End If


finfic:
   
   On Error Resume Next
   Close #nf
   Close #nf2
   Kill vartmp
   Err = 0

End Function

Function f_suplapsav() As String
' recopie lapreso origine
' et suppression lapreso sav

   Dim varfic As String
   Dim nf As Integer
   Dim varstr As String

   f_suplapsav = ""
   
   varfic = "c:\vb\lapreso"
   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic & ".sav" For Input As #nf
   If Err = 0 Then
      Close #nf
      On Error Resume Next
      FileCopy varfic & ".sav", varfic
      If Err = 0 Then
         Kill varfic & ".sav"
      End If
      Err = 0
   End If
   Close #nf
   On Error GoTo 0

   varfic = "c:\se"
   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic & ".sav" For Input As #nf
   If Err = 0 Then
      Close #nf
      On Error Resume Next
      FileCopy varfic & ".sav", varfic
      If Err = 0 Then
         Kill varfic & ".sav"
      End If
      Err = 0
   End If
   Close #nf
   On Error GoTo 0

      'jld20021118 aider
      If lapmmi = 2 Then
         varfic = "c:\numerow"
      Else
         varfic = "c:\essai2"
      End If

   nf = FreeFile
   On Error Resume Next
   Close #nf: Open varfic & ".sav" For Input As #nf
   If Err = 0 Then
      Close #nf
      On Error Resume Next
      FileCopy varfic & ".sav", varfic
      If Err = 0 Then
         Kill varfic & ".sav"
      End If
      Err = 0
   Else
      Err = 0
   End If
   Close #nf
   On Error GoTo 0

    'jld20040105 attention laproot non renseigné
    'jld20040413 masqnorm multienvironnement yyy
   If laproot = "" Then
    varfic = "c:\" & "lapvar.txt"
   Else
    varfic = laproot & "lapvar.txt"
   End If
   On Error Resume Next
   Kill varfic
   
   'jld20040107
   On Error Resume Next
   Kill "c:\vb\lapvar.txt"
   
   On Error Resume Next
   nf = FreeFile
   On Error Resume Next
   'jld20040105
   If laptmp = "" Then
    varfic = "c:\laptmp\" & "lapvar.txt"
   Else
      varfic = laptmp & "lapvar.txt"
   End If
   
   Close #nf: Open varfic For Output As #nf
   If Err = 0 Then
      varstr = Date$ & varsep & Time$ & varsep & laplog & varsep & lappas & varsep & laparg
      varstr = f_cry2(varstr)
      Print #nf, varstr
   End If
   Close #nf
      
   On Error GoTo 0
   
   f_suplapsav = "OK"

End Function

Function f_ntod(varlon1 As Long, varint1 As Integer, varstr1 As String)
'jld20040520 conversion numéro de jour en date
'varlon1 = numéro de jour depuis 01010000
'varint1 = longueur date 6 ou 8
'varstr1 = séparateur

   Dim l As Long
   Dim varlon As Long
   Dim vardat As String
   Dim vartyp As Integer
   Dim varsepdat As String
   Dim varvar As Variant
   Dim varstr As String
   
   f_ntod = ""
   
   varlon = varlon1
   vartyp = varint1
   varsepdat = varstr1
   
   varstr = ""
   On Error Resume Next
   varvar = CVDate(varlon1)
   If Err Then
      varstr = ""
   Else
      Select Case vartyp
      Case 6
         varstr = f_temps(Str(varvar), 1)
      Case 8
      varstr = f_temps(Str(varvar), 2)
      Case Else
      varstr = f_temps(Str(varvar), 2)
      End Select
      If varsepdat <> "" Then
         varstr = Left$(varstr, 2) & varsepdat & Mid$(varstr, 3, 2) & varsepdat & Mid$(varstr, 5)
      End If

   End If
   
   f_ntod = varstr
   
End Function
Function f_temps(varstr1 As String, varint1 As Integer) As String
'varstr1 = date au format DDMMYY ou DDMMYYYY
'varint1 = position de la valeur à retourner : si 0 alors toutes
'retour = 1=DDMMYY, 2=DDMMYYYY, 3=YYYYMMDD, 4=DD/MM/YYYY,
'retour = 5=num jour, 6=abrv jour, 7=nom jour, 8=jour du mois
'retour = 9=nbrjou moi, 10=nommoi, num mois
'retour = 12=année sur 2, 13=année sur 4
'retour = 14=semaine
'retour = 15 nbre jour depuis 31121899

   Dim varok As String
   Dim varrep As String
   Dim varvar As Variant
   Dim varvar2 As Variant
   Dim varstr As String
   Dim vardat As String
   Dim vardat6 As String
   Dim vardat8 As String
   Dim vardat8f As String
   Dim vardati As String
   Dim nomjour As String
   Dim abrjour As String
   Dim numjour As String ' 1 = lundi
   Dim nommoi  As String
   Dim nbrjou As String
   Dim varsem As String
   Dim jo As Integer
   Dim dd As Integer
   Dim mm As Integer
   Dim yy As Integer
   Dim py As Integer
   Dim yyyy As Integer
   Dim ad As Long
   Dim ay As Long
   Dim ax As Integer
   Dim varnum As Integer
   Dim varint As Integer
   Dim vartypdat As String


   f_temps = ""

   varnum = varint1

   'jld20040205 format date entree
   'vardat = Trim$(varstr1)
   
   vardat = Trim$(f_champ(varstr1, ":", 1))
   vartypdat = ""
   vartypdat = Trim$(UCase$(f_champ(varstr1, ":", 2)))


   vardat = Trim$(varstr1)
   
   'jld20020522 attention : détection du point comme séparateur décimal : numérique ?
   'varvar2 = Mid$(vardat, 3, 1)
   'If Not IsNumeric(varvar2) Or varvar2 = "." Then
   '   varstr = varvar2
   '   vardat = f_remplace(vardat, varstr, "", "T", 1)
   'End If

   varstr = ""
   varstr = Mid$(vardat, 3, 1)
   If varstr <> "0" And varstr <> "1" And varstr <> "2" And varstr <> "3" And varstr <> "4" And varstr <> "5" And varstr <> "6" And varstr <> "7" And varstr <> "8" And varstr <> "9" Then
      vardat = f_remplace(vardat, varstr, "", "T", 1)
   End If

   varvar = vardat
   
   If Not IsNumeric(varvar) Then
      Exit Function
   End If
   If Len(vardat) <> 6 And Len(vardat) <> 8 Then
      Exit Function
   End If

   'jld20040205 typdat
   varstr = ""
   varstr = vardat
   Select Case vartypdat
   Case "6"
      vardat = varstr
   Case "8"
      vardat = varstr
   Case "I6"
      varstr = Right$(varstr, 2) & Mid$(varstr, 3, 2) & Left$(varstr, 2)
      vardat = varstr
   Case "I8"
      varstr = Right$(varstr, 2) & Mid$(varstr, 5, 2) & Left$(varstr, 4)
      vardat = varstr
   Case "DI6"
      varstr = Right$(varstr, 2) & Mid$(varstr, 3, 2) & Left$(varstr, 2)
      vardat = varstr
   Case "DI8"
      varstr = Right$(varstr, 2) & Mid$(varstr, 5, 2) & Left$(varstr, 4)
      vardat = varstr
   Case Else
      vardat = varstr
   End Select

   Select Case Len(vardat)
   Case 6
       yy = Val(Right$(vardat, 2))
       mm = Val(Mid$(vardat, 3, 2))
       dd = Val(Left$(vardat, 2))
       If yy > 30 Then
           py = 1900
       Else
           py = 2000
       End If
       yyyy = py + yy
   
   Case 8
         yyyy = Val(Right$(vardat, 4))
           yy = Val(Right$(vardat, 2))
           py = Val(Mid$(vardat, 5, 2) + "00")
           mm = Val(Mid$(vardat, 3, 2))
           dd = Val(Left$(vardat, 2))
       
   Case Else
      Exit Function
   End Select
   
   'vardat8 = Left$(vardat, 4) + Trim$(Str$(py + Val(Right$(vardat, 2))))
   vardat6 = Right$("00" & Trim$(Str$(dd)), 2) & Right$("00" & Trim$(Str$(mm)), 2) & Right$("00" & Trim$(Str$(yy)), 2)
   vardat8 = Right$("00" & Trim$(Str$(dd)), 2) & Right$("00" & Trim$(Str$(mm)), 2) & Right$("0000" & Trim$(Str$(yyyy)), 4)
   vardat8f = Right$("00" & Trim$(Str$(dd)), 2) & "/" & Right$("00" & Trim$(Str$(mm)), 2) & "/" & Right$("0000" & Trim$(Str$(yyyy)), 4)
   vardati = Right$("0000" & Trim$(Str$(yyyy)), 4) & Right$("00" & Trim$(Str$(mm)), 2) & Right$("00" & Trim$(Str$(dd)), 2)
   
   ad = dd
   ay = yyyy
   ad = ad + ay * 365
   
   If mm >= 3 Then
           ad = ad - Int(mm * 0.4 + 2.3)
           ay = ay + 1
   End If
   ad = ad + Int(mm * 31 + (ay - 1) / 4)   'ass
   ax = ad - Int(ad / 7) * 7               'aw
   Select Case ax
       Case 0
           nomjour = "Mardi"
           numjour = "2"
           abrjour = "M"
       Case 1
           nomjour = "Mercredi"
           numjour = "3"
           abrjour = "m"
       Case 2
           nomjour = "Jeudi"
           numjour = "4"
           abrjour = "J"
       Case 3
           nomjour = "Vendredi"
           numjour = "5"
           abrjour = "V"
       Case 4
           nomjour = "Samedi"
           numjour = "6"
           abrjour = "S"
       Case 5
           nomjour = "Dimanche"
           numjour = "7"
           abrjour = "D"
       Case 6
           nomjour = "Lundi"
           numjour = "1"
           abrjour = "L"
   End Select
   
   Select Case mm
       Case 1
            nommoi = "Janvier"
            nbrjou = "31"
       Case 2
            nommoi = "Février"
           If yy Mod 4 <> 0 Then
               nbrjou = "28"
           Else
               nbrjou = "29"
           End If
       Case 3
            nommoi = "Mars"
           nbrjou = "31"
       Case 4
            nommoi = "Avril"
           nbrjou = "30"
       Case 5
            nommoi = "Mai"
           nbrjou = "31"
       Case 6
            nommoi = "Juin"
           nbrjou = "30"
       Case 7
            nommoi = "Juillet"
           nbrjou = "31"
       Case 8
            nommoi = "Août"
           nbrjou = "31"
       Case 9
            nommoi = "Septembre"
           nbrjou = "30"
       Case 10
            nommoi = "Octobre"
           nbrjou = "31"
       Case 11
            nommoi = "Novembre"
           nbrjou = "30"
       Case 12
            nommoi = "Décembre"
           nbrjou = "31"
       Case Else
            nommoi = "***"
           nbrjou = "0"

   End Select
   
   'jld20020105 calcul semaine
   'attention erreur 13 si date fausse   err
   On Error Resume Next
   varsem = ""
   varvar = CVDate("01/01/" & Right$(vardat8, 4))
   varvar2 = CVDate(vardat8f)
   If Err = 0 And Len(varvar) = 10 And Len(varvar2) = 10 Then
      varvar2 = varvar2 + 7 - Val(numjour)
      varint = 0
      Do While True
         varint = varint + 1
         varvar2 = varvar2 - 7
         If varvar2 < varvar Then
            varsem = Trim$(Str$(varint))
            Exit Do
         End If
      Loop
   Else
      varsem = ""
   End If
   Err = 0
   On Error GoTo 0
   
   '6,8,i,/
   varstr = vardat6                               '1  DDMMYY
   varstr = varstr & varsep & vardat8             '2  DDMMYYYY
   varstr = varstr & varsep & vardati             '3  YYYYMMDD
   varstr = varstr & varsep & vardat8f            '4  DD/MM/YYYY
   '1,l,lundi,jj
   varstr = varstr & varsep & numjour             '5  1-7 1=lundi
   varstr = varstr & varsep & abrjour             '6  L M M J V S D
   varstr = varstr & varsep & nomjour             '7  Lundi Mardi
   varstr = varstr & varsep & Left$(vardat6, 2)   '8  01 02 03
   '31,mai
   varstr = varstr & varsep & nbrjou              '9  28 29 30 31
   varstr = varstr & varsep & nommoi              '10 Janvier Février
   varstr = varstr & varsep & Mid$(vardat8, 3, 2) '11 numéro mois          '10 Janvier Février
   'aa,aaaa
   varstr = varstr & varsep & Right$(vardat8, 2)  '12 99
   varstr = varstr & varsep & Right$(vardat8, 4)  '13 1999
   'numéro de semaine : attention semaine 53 le 1er janvier
   varstr = varstr & varsep & varsem 'jld20020105 '14 52
   'nbr jour
   'jld20040520 calcul numéro de jour avec VB
   'varstr = varstr & varsep & Str$(ad) 'jld20020105 '15   731262 = 01012002
   
   'jld20040528 test validité de la date
   On Error Resume Next
   varvar = CVDate(vardat8f)
   If Err Then
      'jld20040816 remise à zéro err si err 13 : 00 dans la date (jour mois) et suppression sortie
      Err = 0

      'Exit Function
   Else
      varstr = varstr & varsep & Trim$(Str$(CLng(varvar))) 'jld20020105 '15   731262 = 01012002
   End If
   


         varok = "O"
         varrep = Left$(vardat8, 2)
         'test jour
         If Val(varrep) > Val(f_champ(varstr, varsep, 9)) Then varok = "N"
         'test mois
         varrep = Mid$(vardat8, 3, 2)
         If Val(varrep) > 12 Then varok = "N"
         'test année
         varrep = Right$(vardat8, 4)
         If Val(varrep) < 1000 Or Val(varrep) > 2999 Then varok = "N"
         If varok = "N" Then
            varstr = ""
         End If

   If varnum = 0 Then
      f_temps = varstr
   Else
      f_temps = Trim$(f_champ(varstr, varsep, varnum))
   End If

   'f_temps = varstr

End Function

Function f_tempsplus(varstr1 As String, varstr2 As String, varlon1 As Long) As String
'opération arithmetique sur date
'varstr1 = date départ sur 8
'varstr2 = opérateur
'varlon1 = valeur
'retour = date mise à jour  sur 8

   Dim varvar As Variant
   Dim vardatdeb As String
   Dim vardatfin As String
   Dim vardat As String
   Dim varstr As String
   Dim varope As String
   Dim varlon As Long
   Dim mois As Integer
   Dim i As Long
   Dim j As Integer
   Dim m As Integer
   Dim a As Integer

   f_tempsplus = ""

   vardatdeb = f_temps(varstr1, 2)
   varope = Trim$(varstr2)
   varlon = varlon1

   'jld20020621
   If varlon = 0 Then
      f_tempsplus = vardatdeb
   End If

   If Len(vardatdeb) <> 8 Then Exit Function
   If Len(varope) <> 1 Then Exit Function
   If varlon < 1 Then Exit Function

   varvar = vardatdeb
   If Not IsNumeric(varvar) Then Exit Function

   vardat = f_temps(vardatdeb, 2)
   If Trim$(vardat) = "" Then Exit Function

   Select Case varope
   Case "+"
      For i = 1 To varlon
          
          j = Val(Left$(vardat, 2))
          m = Val(Mid$(vardat, 3, 2))
          a = Val(Right$(vardat, 4))
          mois = Val(f_temps(vardat, 9))
   
          If j < mois Then
            j = j + 1
          Else
            If m < 12 Then
               j = 1
               m = m + 1
            Else
               j = 1
               m = 1
               a = a + 1
            End If
          End If
   
         vardat = Right$("00" & Trim$(Str$(j)), 2) & Right$("00" & Trim$(Str$(m)), 2) & Right$("0000" & Trim$(Str$(a)), 4)
   
      Next i

   Case "-"
      For i = varlon To 1 Step -1
          
          j = Val(Left$(vardat, 2))
          m = Val(Mid$(vardat, 3, 2))
          a = Val(Right$(vardat, 4))
          mois = Val(f_temps(vardat, 9))
   
          If j > 1 Then
            j = j - 1
          Else
            If m > 1 Then
               j = 1
               m = m - 1
               vardat = Right$("00" & Trim$(Str$(j)), 2) & Right$("00" & Trim$(Str$(m)), 2) & Right$("0000" & Trim$(Str$(a)), 4)
               j = Val(f_temps(vardat, 9))
            Else
               j = 1
               m = 12
               a = a - 1
               vardat = Right$("00" & Trim$(Str$(j)), 2) & Right$("00" & Trim$(Str$(m)), 2) & Right$("0000" & Trim$(Str$(a)), 4)
               j = Val(f_temps(vardat, 9))
            End If
          End If
   
         vardat = Right$("00" & Trim$(Str$(j)), 2) & Right$("00" & Trim$(Str$(m)), 2) & Right$("0000" & Trim$(Str$(a)), 4)
   
          
      Next i

   Case Else
      Exit Function

   End Select

   f_tempsplus = vardat

End Function

Function f_tu(vartmp As String) As String
'FRED20030611

   f_tu = Trim$(UCase$(vartmp))

End Function

Function f_val1(varint1 As Integer, varstr1 As String, varint2 As Integer, varint3 As Integer, varstr2 As String) As String
'jld20020702
'descri=recherche dans tableau tabdes1
'varint1= colonne de recherche
'varstr1= code à chercher
'varint2= colonne de valeur à retourner
'varint3= numéro de sous valeur
'varstr2= séparateur de sous valeur
'retour= valeur cherchée

   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim varcol1 As Integer
   Dim varval1 As String
   Dim varcol2 As Integer
   Dim varssv2 As Integer
   Dim varsep2 As String
   Dim varval2 As String
   Dim i As Integer

   f_val1 = "***"

   On Error Resume Next
   If UBound(tabdes1, 2) < 1 Then
      MsgBox "TABLEAU DE PARAMETRES 1 NON REMPLIT"
      Exit Function
   End If

   On Error GoTo 0

   varcol1 = varint1
   varval1 = Trim$(UCase$(varstr1))
   varcol2 = varint2
   varssv2 = varint3
   varsep2 = varstr2
   
   varval2 = ""
   varstr = ""
   varrep = ""
   varint = 0

   For i = 1 To UBound(tabdes1, 2)
      varrep = Trim$(UCase$(tabdes1(varcol1, i)))
      If varrep = varval1 Then
         varval2 = tabdes1(varcol2, i)
         Exit For
      End If
   Next i

   If varssv2 > 0 Then
      If varsep2 <> "" Then
         varval2 = Trim$(f_champ(varval2, varsep2, varssv2))
      End If
   End If

   f_val1 = varval2


End Function

Function f_var(varstr1 As String) As String
'formatage variables

'jld20040109 evaluation de plusieurs variables séparées par des +
'lapdos+chr$(61)+"998"
'[@CONSEXT;1;A]
'=====================

   Dim varstr As String
   Dim varres As String
   Dim varval As String
   Dim varchp As Integer

   varval = varstr1
   varstr = varstr1

   f_var = varval

   Select Case Left$(Trim$(UCase$(varval)), 2)
   Case "[@"
      varval = Trim(varval)
      varstr = Mid$(Trim$(varval), 3)
      If Right$(varstr, 1) = "]" Then
         varstr = Left$(varstr, Len(varstr) - 1)
      End If
      'If (lapdos <> "" And lapdos <> "0") Or (Trim$(UCase$(f_champ(varstr, ";", 5))) = "F") Then
         varstr = f_valchp(Trim$(f_champ(varstr, ";", 1)), Trim$(f_champ(varstr, ";", 2)), Trim$(f_champ(varstr, ";", 3)), Trim$(f_champ(varstr, ";", 4)), Trim$(f_champ(varstr, ";", 5)), Trim$(f_champ(varstr, ";", 6)), Trim$(f_champ(varstr, ";", 7)))
      'Else
      '   'variable indépendante du dossier
      '   varres = Trim$(UCase$(f_champ(varstr, ";", 1)))
      '   Select Case varres
      '   Case "LAPSER", "LAPDAT", "LAPHEU"
      '      varstr = f_valchp(Trim$(f_champ(varstr, ";", 1)), Trim$(f_champ(varstr, ";", 2)), Trim$(f_champ(varstr, ";", 3)), Trim$(f_champ(varstr, ";", 4)), Trim$(f_champ(varstr, ";", 5)), Trim$(f_champ(varstr, ";", 6)), Trim$(f_champ(varstr, ";", 7)))
      '   Case Else
      '   End Select
      'End If
   End Select
   varval = varstr

   'jld20040607 choix colonne agenda à prendre dans laptmp\agdsel.txt
   'AGD,date8,heure,lapnom,lappre,lapnai,lapdos,,,,COL:1,COL:2,...
   'décallage de 10
   
   Select Case Left$(Trim$(UCase$(varval)), 4)
   Case "COL:"
      varchp = Val(f_champ(varval, ":", 2))
      If varchp > 0 Then
'descri = recherche verticale dans un fichier
'varstr1 = adresse fichier  par défaut dans le dic
'varstr2 = séparateur fichier
'varstr3 = valeur recherchée cle 1
'varint1 = colonne de recherche cle 1
'varstr4 = valeur recherchée cle 2
'varint2 = colonne de recherche cle 2
'varvar1 = colonnes à retourner: 5+8
'varint4 = 0-1 : test majuscules minuscule, 1 : indiférent
'varint4 = 2 : récupère le dernier enregistrement
'retour = valeur trouvée

         varstr = Trim$(f_rechv2(laptmp & "AGENDA\agdsel.txt", varsep, lapnom, 4, lappre, 5, varchp, 1))
      End If
   End Select
   varval = varstr

   Select Case Left$(Trim$(UCase$(varval)), 4)
   Case "CHP:"
      varchp = Val(f_champ(varval, ":", 2))
      If varchp > 0 Then
         varstr = donnees(9, varchp)
      End If
   End Select
   varval = varstr

   'jld20040107 interpretation code ascii caractere
   Select Case Left$(Trim$(UCase$(varval)), 5)
   Case "CHR$("
      varres = f_champ(varval, "(", 2)
      varres = f_champ(varres, ")", 1)
      varchp = Val(varres)
      If varchp > 0 And varchp <= 255 Then
         varstr = Chr$(varchp)
      End If
   End Select
   varval = varstr

   'jld20040107 interpretation "texte"
   Select Case Left$(Trim$(UCase$(varval)), 1)
   Case """"
      varstr = f_champ(varval, """", 2)
      GoTo suitevar
   End Select
   varval = varstr

   Select Case Trim$(UCase$(varval))
   'jld20030729
   Case "LAPDAT"
      varstr = Format$(Now, "DDMMYYYY")
   Case "LAPHEU"
      varstr = Time$
   Case "LAPLOG"
      varstr = laplog
   Case "LAPPAS"
      varstr = lappas
   Case "LAPDOSREP"
      varstr = lapdosrep
   Case "LAPDOSHOP"
      varstr = lapdoshop
   Case "VARSEP"
      varstr = varsep
   Case "VIRGULE"
      varstr = ","
   Case "POINT VIRGULE"
      varstr = ";"
   Case "DEUX POINTS"
      varstr = ":"
   'jld20030922
   Case "LAPSER"
      varstr = lapser
   Case "LAPDOS"
      varstr = lapdos
   Case "LAPNOM"
      varstr = lapnom
   Case "LAPPRE"
      varstr = lappre
   Case "LAPNAI"
      varstr = lapnai
   Case "LAPSEX"
      varstr = lapsex
   Case "LAPNNA"
      varstr = lapnna
   Case "LAPAGE"
      varstr = lapage
   Case "LAPIDE"
      'jld20040916
      varstr = lapide
   Case "LAPMET"
      varstr = lapmet
   Case "LAPNIV"
      varstr = lapniv

   End Select

   'jld20031023 remise en routre modif disparue : variable*
   If InStr(1, varstr, "LAPDAT", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPDAT", Format$(Now, "DDMMYYYY"), "T", 1)
   End If
   If InStr(1, varstr, "LAPHEU", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPHEU", Time$, "T", 1)
   End If
   If InStr(1, varstr, "LAPLOG", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPLOG", laplog, "T", 1)
   End If
   If InStr(1, varstr, "LAPPAS", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPPAS", lappas, "T", 1)
   End If
   If InStr(1, varstr, "LAPDOSREP", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPDOSREP", lapdosrep, "T", 1)
   End If
   If InStr(1, varstr, "LAPDOSHOP", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPDOSHOP", lapdoshop, "T", 1)
   End If
   If InStr(1, varstr, "VARSEP", 1) > 0 Then
      varstr = f_remplace(varstr, "VARSEP", varsep, "T", 1)
   End If
   If InStr(1, varstr, "VIRGULE", 1) > 0 Then
      varstr = f_remplace(varstr, "VIRGULE", ",", "T", 1)
   End If
   If InStr(1, varstr, "POINT VIRGULE", 1) > 0 Then
      varstr = f_remplace(varstr, "POINT VIRGULE", ";", "T", 1)
   End If
   If InStr(1, varstr, "DEUX POINTS", 1) > 0 Then
      varstr = f_remplace(varstr, "DEUX POINTS", ":", "T", 1)
   End If
   If InStr(1, varstr, "LAPSER", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPSER", lapser, "T", 1)
   End If
   If InStr(1, varstr, "LAPDOS", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPDOS", lapdos, "T", 1)
   End If
   If InStr(1, varstr, "LAPNOM", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPNOM", lapnom, "T", 1)
   End If
   If InStr(1, varstr, "LAPPRE", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPPRE", lappre, "T", 1)
   End If
   If InStr(1, varstr, "LAPNAI", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPNAI", lapnai, "T", 1)
   End If
   If InStr(1, varstr, "LAPSEX", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPSEX", lapsex, "T", 1)
   End If
   If InStr(1, varstr, "LAPNNA", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPNNA", lapnna, "T", 1)
   End If
   If InStr(1, varstr, "LAPAGE", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPAGE", lapage, "T", 1)
   End If
   If InStr(1, varstr, "LAPIDE", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPIDE", lapide, "T", 1)
   End If
   If InStr(1, varstr, "LAPMET", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPMET", lapmet, "T", 1)
   End If
   If InStr(1, varstr, "LAPNIV", 1) > 0 Then
      varstr = f_remplace(varstr, "LAPNIV", lapniv, "T", 1)
   End If
   
   'Select Case Left$(Trim$(UCase$(varval)), 4)
   'Case "CHP:"
   '   varchp = Val(f_champ(varval, ":", 2))
   '   If varchp > 0 Then
   '      varstr = donnees(9, varchp)
   '   End If
   'End Select

suitevar:

   f_var = varstr

End Function

Function fg_cpfic(varstr1 As String, varstr2 As String, varstr3 As String) As String
'Cette fonction copie le contenu d'un fichier dans un autre
'
'varstr1 = fichier donnant
'varstr2 = fichier recevant
'varstr3 = "OUT" pour output (ecrasement du receveur)
'varstr3 = "APP" pour append (rajout au receveur)
'varstr3 = "APP@@contenu@@numero" : remplace la ligne numero n par le contenu fourni
'varstr3 = "APP^^contenu^^numero" : ajout contenu en début de ligne numero :si numero = 0 alors toutes ls lignes
'varstr3 = "APP^R^contenu^R^numero" : ajout contenu en début de ligne numero ssi chaine non presente :si numero = 0 alors toutes ls lignes
'varstr3 = "APP^C^contenu^C^numero" : remplace champ 1 d'une ligne par contenu :si numero = 0 alors toutes ls lignes
'varstr3 = "APP$$contenu$$numero" : ajout contenu en fin de ligne numero :si numero = 0 alors toutes ls lignes
'varstr3 = "APP$R$contenu$R$numero" : remplace contenu en fin de ligne numero :si numero = 0 alors toutes ls lignes
'varstr3 = "APP$RR$contenu#remplacement$RR$numero" : supprime contenu en fin de ligne numero :si numero = 0 alors toutes ls lignes
'varstr3 = "APP^CC^contenu#sep#valeur#chaine^CC^numero" : chaine remplace la ligne numero si le champ valeur = contenu, sep = separateur de champ, marche pour toute les lignes si numero = 0
'varstr3 = "APP::N" : pas de message si erreur de lecture fichier source
'varstr3 = "APP&&VALEUR&&CHAMP : ajoute la ligne qui contient VALEUR dans le champ nr : CHAMP
'
'Retour = nombre de lignes si tout a ete fait correctement, sinon vide
'

'FRED20021106

   Dim varchp As Integer 'jld20040903
   Dim varval As String  'jld20040903
   Dim varmes As String
   Dim varcod As String
   Dim ficdon As String
   Dim ficrec As String
   Dim fictyp As String
   Dim lnum As Long
   Dim nvligne As String
   Dim lect As String
   Dim nf As Integer
   Dim nf2 As Integer
   Dim k As Long
   Dim x As Integer
   Dim xx As Integer
   Dim tmp1 As String
   Dim tmp2 As String
   Dim tmp3 As String
   Dim tmp4 As String

   Dim tablig() As String
   Dim tabnum() As Long
   
   'fred20020617
   On Error Resume Next

   ficdon = varstr1
   ficrec = varstr2
   
   'jld20020910
   'If InStr(varstr3, "@@") <> 0 Then
   '   fictyp = UCase$(Trim$(f_champ(varstr3, "@@", 1)))
   '   nvligne = Trim$(f_champ(varstr3, "@@", 2))
   '   lnum = CLng(Trim$(f_champ(varstr3, "@@", 3)))
   'Else
   '   fictyp = UCase$(Trim$(varstr3))
   'End If
   'err = 0
   
   varcod = ""
   fictyp = ""
   nvligne = ""
   lnum = 0
   ReDim tablig(1)
   ReDim tabnum(1)

   'jld20020921
   varmes = ""
   If InStr(varstr3, "::") <> 0 Then
      varmes = Trim$(UCase$(f_champ(varstr3, "::", 2)))
      varstr3 = Trim$(UCase$(f_champ(varstr3, "::", 1)))
   End If

   lect = ""
   x = 1
   If InStr(varstr3, "@@") <> 0 Then
      varcod = "@@"
      fictyp = UCase$(Trim$(f_champ(varstr3, "@@", 1)))
      nvligne = Trim$(f_champ(varstr3, "@@", 2))
      lnum = CLng(Trim$(f_champ(varstr3, "@@", 3)))
      lect = "@@"
      Do While lect <> ""
         lect = Trim$(f_champ(varstr3, "@@", x * 2))
         If lect <> "" Then
            tablig(x) = lect
            tabnum(x) = CLng(Trim$(f_champ(varstr3, "@@", x * 2 + 1)))
            x = x + 1
            ReDim Preserve tablig(x)
            ReDim Preserve tabnum(x)
         End If
      Loop
      x = x - 1
      ReDim Preserve tablig(x)
      ReDim Preserve tabnum(x)
   End If

   If InStr(varstr3, "^^") <> 0 Then
      varcod = "^^"
      fictyp = UCase$(Trim$(f_champ(varstr3, "^^", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "^^", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "^^", 3)))
   End If

   If InStr(varstr3, "^R^") <> 0 Then
      varcod = "^R^"
      fictyp = UCase$(Trim$(f_champ(varstr3, "^R^", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "^R^", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "^R^", 3)))
   End If

   If InStr(varstr3, "^C^") <> 0 Then
      varcod = "^C^"
      fictyp = UCase$(Trim$(f_champ(varstr3, "^C^", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "^C^", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "^C^", 3)))
   End If

   If InStr(varstr3, "$$") <> 0 Then
      varcod = "$$"
      fictyp = UCase$(Trim$(f_champ(varstr3, "$$", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "$$", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "$$", 3)))
   End If

   If InStr(varstr3, "$R$") <> 0 Then
      varcod = "$R$"
      fictyp = UCase$(Trim$(f_champ(varstr3, "$R$", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "$R$", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "$R$", 3)))
   End If

   If InStr(varstr3, "$RR$") <> 0 Then
      varcod = "$RR$"
      fictyp = UCase$(Trim$(f_champ(varstr3, "$RR$", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "$RR$", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "$RR$", 3)))
   End If

   If InStr(varstr3, "^CC^") <> 0 Then
      varcod = "^CC^"
      fictyp = UCase$(Trim$(f_champ(varstr3, "^CC^", 1)))
      'on garde ajout avec espaces
      nvligne = f_champ(varstr3, "^CC^", 2)
      lnum = CLng(Trim$(f_champ(varstr3, "^CC^", 3)))
   End If

   'jld20040903 choit ligne à ajouter
   If InStr(varstr3, "&&") <> 0 Then
      varcod = "&&"
      fictyp = Trim$(UCase$(f_champ(varstr3, "&&", 1)))
      varval = f_champ(varstr3, "&&", 2)
      varchp = Val(Trim$(f_champ(varstr3, "&&", 3)))
   End If


   If varcod = "" Then
      fictyp = UCase$(Trim$(varstr3))
   End If
   On Error Resume Next

   nf = FreeFile
   
   'fred20020617
   'nf2 = nf + 1
   fg_cpfic = ""
   k = 0
   
   'fred20020617
   Close #nf

   If fictyp = "" Then fictyp = "OUT"

   On Error Resume Next

   Select Case fictyp
   Case "OUT"
      Open ficrec For Output Lock Read Write As #nf  'jld20031203 : Lock #nf

   Case "APP"
      Open ficrec For Append Lock Read Write As #nf  'jld20031203 : Lock #nf

   Case Else
   End Select

   If Err = 0 Then
    
      'fred20020617
      xx = 1
      nf2 = FreeFile
      Close #nf2
    
      Open ficdon For Input As #nf2
    If Err = 0 Then
       Do While Not EOF(nf2)
          On Error Resume Next
          Line Input #nf2, lect
          If Err Then
            'jld20020921
            If varmes = "" Then
               x = MsgBox("Impossible de lire le fichier " & UCase$(f_basname(ficdon)), 64, "WIN_LAP")
            End If
            Exit Do
          End If
      
          'jld20020910
          Select Case varcod
          Case ""
             Print #nf, lect
          Case "@@"
             'If lnum = k + 1 Then
             '   Print #nf, nvligne
             'Else
             '   Print #nf, lect
             'End If
             x = CLng(tabnum(xx))
             If x = k + 1 And Err = 0 Then
                 Print #nf, tablig(xx)
                 xx = xx + 1
             Else
                 Print #nf, lect
             End If
             Err = 0
      
          Case "^^"
             If lnum = k + 1 Or lnum = 0 Then
                 Print #nf, nvligne; lect
             Else
                 Print #nf, lect
             End If
      
          Case "^R^"
             If lnum = k + 1 Or lnum = 0 Then
               If Left$(lect, Len(nvligne)) <> nvligne Then
                 Print #nf, nvligne; lect
               Else
                 Print #nf, lect
               End If
             Else
               If Left$(lect, Len(nvligne)) <> nvligne Then
                 Print #nf, lect
               End If
            End If
          
          Case "^C^"
             If lnum = k + 1 Or lnum = 0 Then
               If Trim$(UCase$(lect)) = "" Then
                 Print #nf, nvligne
               Else
                 Print #nf, f_remplace(lect, f_champ(lect, ",", 1), nvligne, "", 0)
               End If
             Else
               Print #nf, lect
             End If
      
          Case "$$"
             If lnum = k + 1 Or lnum = 0 Then
                 Print #nf, lect; nvligne
             Else
                 Print #nf, lect
             End If
             
          Case "$R$"
             If lnum = k + 1 Or lnum = 0 Then
               If Len(lect) > Len(nvligne) Then
                 Print #nf, Left$(lect, Len(lect) - Len(nvligne)); nvligne
               Else
                 Print #nf, nvligne
               End If
             Else
              Print #nf, lect
             End If
          
          Case "$RR$"
             If lnum = k + 1 Or lnum = 0 Then
               tmp1 = f_champ(nvligne, "#", 1)
               tmp2 = f_champ(nvligne, "#", 2)
               If Right$(lect, Len(tmp1)) = tmp1 Then
                 Print #nf, Left$(lect, Len(lect) - Len(tmp1)); tmp2
               Else
                 Print #nf, lect
               End If
             Else
              Print #nf, lect
             End If
        
          Case "^CC^" 'FRED20030506 A finir
             If lnum = k + 1 Or lnum = 0 Then
               tmp1 = Trim$(UCase$(f_champ(nvligne, "#", 1))) 'Contenu a tester
               tmp2 = f_champ(nvligne, "#", 2) 'Separateur de champ
               tmp3 = f_champ(nvligne, "#", 3) 'Numero du champ a tester
               tmp4 = f_champ(nvligne, "#", 4) 'chaine de remplacement
               If Trim$(UCase$(f_champ(lect, tmp2, CInt(tmp3)))) = tmp1 Then
                  Print #nf, tmp4
               Else
                  Print #nf, lect
               End If
             Else
              Print #nf, lect
             End If
            
          Case "&&" 'jld20040903 choix lignes à ajouter
               If Trim$(UCase$(f_champ(lect, varsep, varchp))) = Trim$(UCase$(varval)) Then
                  Print #nf, lect
               End If
            
          Case Else
         
          End Select
      
          k = k + 1
       Loop
       fg_cpfic = CStr(k)

    Else

       'jld20030106 fichier inexistant
       If Err = 53 Then
          fg_cpfic = "0"
       End If
   
       'jld20020921
       If varmes = "" Then
          x = MsgBox("Impossible de lire le fichier " & UCase$(f_basname(ficdon)), 64, "WIN_LAP")
       End If
       Err = 0
    End If
      Close #nf2

   Else
      'jld20020921
      If varmes = "" Then
    x = MsgBox("Impossible de creer le fichier " & UCase$(f_basname(ficrec)), 64, "WIN_LAP")
      End If
      Err = 0
   End If
   
   'jld20031224 Unlock #nf
   Close #nf

   Err = 0

End Function

Function fg_exist(varstr1 As String, varstr2 As String, varstr3 As String) As Integer
' Fred20011207
' Cette fonction controle l'existence d'un champ dans une ligne de champs
'
' varstr1 = ligne de champs
' varstr2 = separateur de champs
' varstr3 = valeur recherchee
'
' Retour : derniere position du champ si trouve, sinon 0

   Dim ligne As String
   Dim fgsep As String
   Dim champ As String
   Dim fgtmp As String
   Dim i As Integer
   Dim x As Integer
   
   ligne = varstr1
   fgsep = varstr2
   champ = varstr3
   fg_exist = 0
   i = 1
   
   'Modif Fred20011213, on peut passer un espace en separateur de champs
   'If Trim$(fgsep) = "" Then fgsep = varsep
   ligne = Trim$(ligne)


   If Not champ = "" Or Not ligne = "" Then
      Err = 0
      i = 1
      x = 1
      'fgtmp = f_champ(ligne, fgsep, 1)
      Do While Not i = 0
         i = InStr(ligne, fgsep)
         fgtmp = f_champ(ligne, fgsep, 1)
         If fgtmp = champ Then
            fg_exist = x
            Exit Do
         End If
         ligne = Mid$(ligne, i + 1)
         x = x + 1
      Loop

   End If

End Function

Sub fg_lectpalm()

   'Implementee Fred20011031 Multijour
   'Cette fonction lit le fichier alisutil.txt a la partie PALM
   'Attention, toutes les variables usitees sont globales

   Dim x As Variant
   Dim y As Variant
   Dim fgtmp1 As String
   Dim fgtmp2 As String
   Dim fgtmp3 As String
   Dim nfg As Integer
   
   ''On determine le jour actuel de la semaine
   x = fg_today()
   
   ''Lecture du nombre de jours pour la sauvegarde config services
   nfg = FreeFile
   Dim fglect As String
   On Error Resume Next
   Open formatrep(lapvb) & "\alisutil.txt" For Input As #nfg
   If Err = 0 Then
    Do While Not EOF(nfg)
       'jld20040816
       On Error Resume Next
       Line Input #nfg, fglect
      If Err Then 'jld20011203
         MsgBox Str(Err) & " ERREUR WHILE lectpalm 1"
         Err = 0
         Exit Do
      End If
       
       If Not Right$(Trim$(fglect), 1) = "'" Then
          
          If Trim$(UCase$(fglect)) = "[MULTI]" Then
        Do While Not EOF(nfg)
               'jld20040816
            On Error Resume Next
            Line Input #nfg, fglect
            If Err Then 'jld20011203
               MsgBox Str(Err) & " ERREUR WHILE lectpalm 1"
               Err = 0
               Exit Do
            End If

           If Not Left$(Trim$(UCase$(fglect)), 1) = "[" Then  'On lit jusqu'au prochain paragraphe
         
            fglect = Trim$(UCase$(fglect))
            x = Left$(fglect, InStr(fglect, "=") - 1) 'Determine le jour
            y = Right$(fglect, Len(fglect) - InStr(fglect, "=")) 'Determine le nombre d'anteriorites a conserver
            
            If y > 0 Then y = y + 1 'Si c'est une journee vaide on tient compte du jour actuel en plus, sinon on ne prend rien
   
            fgnbjour = y
            
            Select Case x
               Case "LUNDI"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 1
                  nbjlu = y
               Case "MARDI"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 1
                  nbjma = y
               Case "MERCREDI"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 1
                  nbjme = y
               Case "JEUDI"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 1
                  nbjje = y
               Case "VENDREDI"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 3
                  nbjve = y
               Case "SAMEDI"
                  If Trim$(CStr(y)) = "" Then y = 0
                  nbjsa = y
               Case "DIMANCHE"
                  If Trim$(CStr(y)) = "" Then y = 0
                  nbjdi = y
               Case "ARCHIVES"
                  If Trim$(CStr(y)) = "" Or y = 0 Then y = 7
                  nbjtot = y
            End Select

           Else
            Exit Do  'On arrete de lire le paragraphe suivant
           End If
        Loop
          End If
       End If
    Loop
   Else
      Err = 0
      MsgBox "Erreur lecture : " & formatrep(lapvb) & "\alisutil.txt"
   End If
   Close #nfg
   Err = 0


End Sub

Function fg_today() As String

'Implementee Fred20011031 Multijour
'Retourne la date systeme au format DDMMYYYY

   Dim x As String
   Dim fgtmp1 As String
   Dim fgtmp2 As String
   Dim fgtmp3 As String

   x = Date$
   fgtmp1 = Mid$(x, 4, 2)
   fgtmp2 = Left$(x, 2)
   fgtmp3 = Right$(x, 4)
   x = fgtmp1 & fgtmp2 & fgtmp3

   fgjour = f_temps(CStr(x), 7)
   fgjour = Trim$(UCase$(fgjour))

   fg_today = x

End Function

Function formatrep(rep As String) As String
'Fonction Fred20011024
'enlève le dernier anti slash de l adresse

   Dim fgtmprep As String

   fgtmprep = rep

   formatrep = fgtmprep


    'Mise en forme de la saisie des repertoires
    If Right(fgtmprep, 1) = "\" Then
         fgtmprep = Left(fgtmprep, Len(fgtmprep) - 1)
    End If
    
    formatrep = fgtmprep  'Retourne le repertoire formate  sans le "\"

End Function

Function mk_dir(varchem As String, varmsgbox As Integer) As String

'Fonction implementee par Fred20011026, creation de repertoire avec sous-repertoire

'Equivalent a MkDir, sauf que ca cree aussi les sous-rep s'ils n'existent pas
'Verifie aussi le chemin a creer au format 8.3 crs sur un lecteur valide
'varchem = chemin a creer
'varmsgbox = avec message a chaque creation (1 oui, 0 non)
'Retourne le chemin cree, et pb = 1 si problemes rencontres

pb = 0
mk_dir = ""

'Memorisation du chemin (profondeur de 200 rep max)
Dim tbchem() As String
ReDim tbchem(200) As String
Dim tmpstr1 As String
Dim tmpstr2 As String
Dim x As Variant
Dim k As Variant
Dim longtab As Variant
Dim origdir As String 'fred20011126

origdir = CurDir$ 'fred20011126

tmpstr1 = varchem
x = InStr(tmpstr1, "\")
k = 0

While Not x = 0
    tmpstr1 = formatrep(tmpstr1)
    x = InStr(tmpstr1, "\")
    If Not x = 0 Then
   tmpstr2 = Left$(tmpstr1, x - 1) 'Rep haut
   tbchem(k) = tmpstr2
   tmpstr1 = Right$(tmpstr1, Len(tmpstr1) - x) 'Rep bas
   k = k + 1
    End If
    If x = 0 Then
   tmpstr2 = tmpstr1 'Rep haut
   tbchem(k) = tmpstr2
   k = k + 1
    End If
Wend
longtab = k


'Verification saisie correcte du lecteur
If Not Len(tbchem(0)) = 2 Then
    MsgBox "Veuiller préciser une lettre de lecteur valide s'il vous plait"
    pb = 1
    Exit Function
End If


''Verification du respect du format 8 crs pour les noms de rep
'k = 0
'While Not k = longtab
'   If Len(tbchem(k)) > 8 Then
'      MsgBox "Veuillez garder des noms de repertoire de 8 caracteres au maximum"
'      pb = 1
'      Exit Function
'   End If
'   k = k + 1
'Wend


'Creation repertoires
k = 0
tmpstr1 = ""
While Not k = longtab
    If Not tbchem(k) = "" Then
   tmpstr1 = tmpstr1 + tbchem(k) + "\"
    End If
    
    On Error Resume Next
    ChDir tmpstr1
    If Err Then
      Err = 0
      tmpstr2 = Left(tmpstr1, Len(tmpstr1) - 1)
      
      If varmsgbox = 1 Then
          x = MsgBox("Le répertoire " + UCase$(tmpstr2) + "\" + " n'existe pas, voulez-vous le créer ?", 36, "WIN_LAP SETUP")
      Else
          x = 6
      End If
      
      If x = 7 Then
          pb = 1
          Exit Function
      End If
   
      If x = 6 Then
       On Error Resume Next
       MkDir tmpstr2
       If Err Then
         Err = 0
      MsgBox "Impossible sur ce lecteur"
      mk_dir = ""
      pb = 1
      Exit Function
       End If
       mk_dir = tmpstr2
   End If

    Else
    mk_dir = varchem
    
    End If
    
    k = k + 1
    
Wend

ChDir origdir 'fred20011126

End Function

Sub modelemsg()

'=================================================
'Debut message avec reponse oui non

   
   'variables.

   Dim DgDef
   Dim msg
   Dim Response
   Dim Title
   
   Title = "Message Demo"
   msg = "ATTENTION "
   msg = msg & "vous allez faire qqe chose : "
   msg = msg & "voulez vous continuer ?"

   'dialogue.

   DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2

   'réponse.
   Response = MsgBox(msg, DgDef, Title)

   'évaluation.
   If Response = IDYES Then
      msg = "choix OUI."
   Else
      msg = "choix NON."
   End If

   MsgBox msg

'fin message avec reponse oui non
'=================================================

End Sub

Sub pos_form(varform As Form)
     'Fction rajoutee par Fred20011022
Dim screenx As Integer
Dim screeny As Integer
Dim formx As Integer
Dim formy As Integer


'Centre une forme Visual Basic au milieu de l'ecran, et diminue sa taille si elle
'depasse de l'ecran
  
    'On verifie si la taille ne dépasse pas, sinon on réduit
    If varform.Width > Screen.Width Then varform.Width = Screen.Width - 100
    If varform.Height > Screen.Height Then varform.Height = Screen.Height - 100

    'On centre la forme
    screenx = Val(Screen.Width / 2)
    screeny = Val(Screen.Height / 2)
    formx = Val(varform.Width / 2)
    formy = Val(varform.Height / 2)
    varform.Top = screeny - formy
    varform.Left = screenx - formx
   
End Sub

Sub pos_obj(varobj As Control, varform As Form, dechor As Long, decver As Long)
'Fction rajoutee par Fred20020404
Dim formx As Integer
Dim formy As Integer
Dim objx As Integer
Dim objy As Integer

'Centre un objet au milieu d'une forme
  
    'On verifie si la taille ne dépasse pas, sinon on réduit
    If varobj.Width > varform.Width Then varobj.Width = varform.Width - 100
    If varobj.Height > varform.Height Then varobj.Height = varform.Height - 100

    'On centre l'objet
    formx = Val(varform.Width / 2)
    formy = Val(varform.Height / 2)
    objx = Val(varobj.Width / 2)
    objy = Val(varobj.Height / 2)
    varobj.Top = formy - objy - decver 'Sinon on a un decalage vertical variable
    varobj.Left = formx - objx - dechor 'Sinon on a un decalage horizontal variable
   
End Sub

Sub s_erreur(varstr1 As String, varvar1 As Variant, varstr2 As String, varstr3 As String)
'gestion erreur dans laptra\laperr.log
'varstr1 = fichier en erreur
'varvar1 = numéro erreur : RW
'varstr2 = commentaire
'varstr3 = ligne en cause

   Dim varcom As String
   Dim varact As String
   Dim varfic As String
   Dim varlig As String
   Dim varstr As String
   Dim vardat As String
   Dim vartim As String
   Dim varerr As String
   Dim nf As Integer

   varfic = Trim$(varstr1)
   varcom = Trim$(varstr2)
   varlig = Trim$(varstr3)
   varstr = CStr(varvar1)
   varerr = Trim$(UCase$(f_champ(varstr, ":", 1)))
   varact = Trim$(UCase$(f_champ(varstr, ":", 2)))

   If Trim$(laptra) = "" Then Exit Sub

   vardat = Date$
   vartim = Time$

   On Error Resume Next
   nf = FreeFile
   Close #nf: Open laptra & "laperr.log" For Append As #nf
   If Err = 0 Then

      varstr = "LAPERR"
      varstr = varstr & varsep & "¥" & vardat
      varstr = varstr & varsep & vartim
      varstr = varstr & varsep & Format$(Now, "DDMMYYYY")
      varstr = varstr & varsep & Time$
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & ""
      
      varstr = varstr & varsep & lapser  'service
      varstr = varstr & varsep & Laphos  'hospit
      varstr = varstr & varsep & lappat  'patient
      varstr = varstr & varsep & lapdoc  'numéro doc
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & varfic  'fichier
      varstr = varstr & varsep & varerr  'erreur
      varstr = varstr & varsep & varact  'action
      varstr = varstr & varsep & lapide  'ide user
      varstr = varstr & varsep & lapniv  'niveau user
      
      varstr = varstr & varsep & varcom
      varstr = varstr & varsep & ""
      varstr = varstr & varsep & ""

      Print #nf, varstr

   End If
   Close #nf

   On Error Resume Next
   nf = FreeFile
   Close #nf: Open laptra & "lapliger.log" For Append As #nf
   If Err = 0 Then

      Print #nf, varlig

   End If
   Close #nf

End Sub

Sub s_listfont()

   Dim i
   ReDim tabstr(0)
   
   On Error Resume Next
   For i = 0 To Printer.FontCount - 1
      'Listfont.AddItem Printer.Fonts(I)
      ReDim Preserve tabstr(i + 1)
      tabstr(i + 1) = Printer.Fonts(i)
      If Err Then Exit Sub
   Next i
   On Error GoTo 0

End Sub

Function f_valchp(varstr1 As String, varstr2 As String, varstr3 As String, varstr4 As String, varstr5 As String, varstr6 As String, varstr7 As String) As String
'DESCRI=retourne la valeur du champ
'varstr1=masque
'varstr2=champ
'varstr3=type
'varstr4=ligne
'varstr5=fichier
'varstr6=separateur
'varstr7=adresse

   Dim vartrv As String
   Dim varche As String
   Dim varchp As Integer
   Dim varlig As String
   Dim varfic As String
   Dim varmsq As String
   Dim varstr As String
   Dim varrep As String
   Dim varint As Integer
   Dim i As Integer
   Dim vartypchp As String
   Dim varnumlig As Integer
   Dim varcptlig As Integer
   Dim vartypfic As String
   Dim varsepfic As String
   Dim varadrfic As String
   Dim nf As Integer
   Dim a$

   f_valchp = ""
      
   varmsq = Trim$(UCase$(varstr1))
   
   varchp = Val(varstr2)
   
   vartypchp = Trim$(UCase$(varstr3))
   If vartypchp = "" And varchp <> 0 Then
      vartypchp = "A" 'alpha ou N numérique
   End If

   varnumlig = Val(varstr4)
   
   vartypfic = Trim$(UCase$(varstr5))
   If vartypfic = "" Then
      vartypfic = "M"
   End If

   varsepfic = varstr6
   If varsepfic = "" Then
      varsepfic = varsep
   End If
   If Trim$(UCase$(varsepfic)) = "VIRGULE" Then
      varsepfic = ","
   End If
   If Trim$(UCase$(varsepfic)) = "POINT VIRGULE" Then
      varsepfic = ";"
   End If
   If Val(varsepfic) = 59 Then
      varsepfic = ";"
   End If
   
   varadrfic = Trim$(UCase$(varstr7))
   varadrfic = f_adr(varadrfic)

   Select Case varmsq
   Case "LAPAGE"
      varrep = lapage
      vartrv = "OK"
   
   Case "LAPNOM"
      varrep = lapnom
      vartrv = "OK"
   
   Case "LAPNNA"
      varrep = lapnna
      vartrv = "OK"
   
   Case "LAPPRE"
      If vartypchp = "P" Then
      varint = f_remplit_str(lappre, " ")
      For i = 1 To varint
         If i = 1 Then
            varrep = UCase$(Left$(tabstr(i), 1)) & LCase$(Mid$(tabstr(i), 2))
         Else
            varrep = varrep & " " & UCase$(Left$(tabstr(i), 1)) & LCase$(Mid$(tabstr(i), 2))
         End If
      Next i
      Else
         varrep = lappre
      End If
      vartrv = "OK"
   
   Case "LAPNAI"
      varrep = lapnai
      vartrv = "OK"
   
   Case "LAPSEX"
      varrep = lapsex
      vartrv = "OK"
   
   Case "LAPDOS"
      varrep = lapdos
      vartrv = "OK"
   
   Case "LAPSER"
      varrep = lapser
      vartrv = "OK"
   
   Case "LAPDAT"
      varrep = Format$(Now, "DDMMYYYY")
      vartrv = "OK"
   
   Case "LAPHEU"
      varrep = Time$
      vartrv = "OK"
   
   Case "LAPIDE"
      varrep = lapide
      vartrv = "OK"
   
   Case "LAPMET"
      varrep = lapmet
      vartrv = "OK"
   
   Case "LAPNIV"
      varrep = lapniv
      vartrv = "OK"
   
   Case Else

      Select Case vartypfic
      Case "M"
      
         'jld20040413 masqnorm multienvironnement
         'varfic = laproot & "interrtf.txt"
         varfic = laproot & lapmsq & "interrtf.txt"
         varsepfic = varsep
         If varchp <> 0 Then varchp = varchp + 3
      
         'jld20011021
         'If Left$(Trim$(varfic), 2) = "\\" Then
         '   varfic = f_remplace(varfic, "\\", "\", "TOUS", 1)
         '   varfic = "\" & varfic
         'Else
         '   varfic = f_remplace(varfic, "\\", "\", "TOUS", 1)
         'End If
         varfic = f_adr(varfic)
         
         Err = 0
         varrep = "" 'jld
         On Error Resume Next
         varrep = Dir(varfic, 0)
      
         varrep = ""
         vartrv = ""
         
         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf
         If Err Then
            Close #nf
            vartrv = ""
            varrep = ""
            Exit Function
         Else
      
            vartrv = ""
            varrep = ""
            
            Do While Not EOF(nf)
               Line Input #nf, a$
               varlig = a$
               
               If Len(varlig) > 0 And InStr(1, varlig, varsepfic) > 0 Then
                     If UCase$(f_champ(varlig, varsepfic, 1)) = varmsq Then
                        varrep = Trim$(f_champ(varlig, varsepfic, varchp))
                     End If
               End If
            Loop
            Close #nf
            
         End If 'err
      
      Case "D"
         MsgBox "Type fichier D : en cours de developpement"
      Case "F"
         varrep = ""
         varcptlig = 0
         varfic = varadrfic & varmsq
         varfic = f_adr(varfic)

         nf = FreeFile
         On Error Resume Next
         Close #nf: Open varfic For Input As #nf
         If Err Then
            MsgBox Str(Err) & " ERREUR LECTURE fichier : " & varmsq
         Else
            Do While Not EOF(nf)
               Line Input #nf, a$
               varcptlig = varcptlig + 1
               If varcptlig = varnumlig Then
                  varrep = f_champ(a$, varsepfic, varchp)
               End If
            Loop
         End If
         Close #nf

      Case Else
         MsgBox "Type de fichier non reconnu"
      End Select

   End Select


   varrep = f_typchp(varrep, vartypchp)
   f_valchp = varrep

End Function
Function f_typchp(varstr1 As String, varstr2 As String) As String
'varstr1 = valeur à formater
'varstr2 = type champ
'"A"    'alpha
'"AA"   'année sur 2 caractères
'"AAAA" 'année sur 4 caractères
'"AAAAMMJJ" 'date inverse               'jld20030710
'"DI","DI6",DI8" jld20040207 date inversée 6 ou 8 à remettre à l endroit
'"D"    'date : DD/MM/YYYY
'"E"    'date : DD mois YYYY
'"F"    'date : Jour DD mois YYYY
'"G"    'date : Jour
'"H"    'heure HH"H"MM 12H34
'"JJ"   'jour sur 2 caractères
'"L"    'multiligne
'"M"    'numérique avec ___ si rien
'"MM"   'mois sur 2 caractères
'"MO"   'mois libelle
'"N"    'numérique
'"P"    'prénom : première lettre en majuscule Jean-Louis
'"R"    'alpha remplacement laprtfrp.txt
'"T"    'heure HH:MM
'"HHMMSS"
'LTOD   'conversion 1/4 1/2 3/4 en décimale
'"^"    'alpha avec blanc transformé en ^ : dicom  'jld20030710
':car:taille:G/D = remplissage et taille du champ
':debut:taille:P = partie de champ (si taille = * alors le reste)

   Dim varcar As String
   Dim vartaille As String
   Dim vardg As String
   Dim varvar As Variant
   Dim varrep As String
   Dim varstr As String
   Dim vartypchp As String
   Dim varint As Integer
   Dim nf As Integer
   Dim i As Integer
   Dim a$

   varrep = varstr1
   'vartypchp = varstr2
   vartypchp = Trim$(UCase$(f_champ(varstr2, ":", 1)))
   varcar = f_champ(varstr2, ":", 2)
   vartaille = Trim$(f_champ(varstr2, ":", 3))
   vardg = Trim$(UCase$(f_champ(varstr2, ":", 4)))

   f_typchp = varrep

      Select Case vartypchp
      Case "^"  'alpha blanc=^
         varrep = f_remplace(varrep, " ", "^", "T", 1)
      
      Case "A"  'alpha
         varrep = varrep
      
      Case "AA" 'année sur 2 caractères
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = Right$(varrep, 2)

      Case "AAAA" 'année sur 4 caractères
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = Right$(varrep, 4)
      Case "HHMMSS" 'heure
         If Trim$(varrep) = "" Then varrep = "******"
         Select Case Len(Trim$(varrep))
         Case 8 'HHxMMxSS"
            varrep = Trim$(f_champ(varrep, ":", 1)) & Trim$(f_champ(varrep, ":", 2)) & Trim$(f_champ(varrep, ":", 3))
         Case 5 'HHxMM
            varrep = Left$(varrep, 2) & Right$(varrep, 2) & "00"
         Case 4 'HHMM
            varvar = varrep
            If IsNumeric(varvar) Then
               varrep = Left$(varrep, 2) & Right$(varrep, 2) & "00"
            End If
         Case Else
         End Select
      Case "YYYYMMDD" 'date inverse 8
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = f_temps(varrep, 3)
      Case "YYMMDD" 'date inverse 6
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = f_temps(varrep, 3)
         varrep = Mid$(varrep, 3)
      Case "DDMMYY" 'date 6
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = f_temps(varrep, 1)
      
            'jld20040207 date inversée 6 ou 8 à remettre à l endroit
      Case "DI", "DI6", "DI8" 'date : YYYYMMDD -> DDMMYYYY
         If Trim$(varrep) = "" Then varrep = "********"
         Select Case Len(varrep)
         Case 6
            varrep = Right$(varrep, 2) & Mid$(varrep, 3, 2) & Left$(varrep, 2)
         Case 8
            varrep = Right$(varrep, 2) & Mid$(varrep, 5, 2) & Left$(varrep, 4)
         End Select

      Case "D"   'date : DD/MM/YYYY
         'jld20041102
         varrep = f_temps(varrep, 4)
         If Trim$(varrep) = "" Then varrep = "**/**/****"
         
         'If Trim$(varrep) = "" Then varrep = "********"
         'If Len(varrep) = 8 Then
         '   varrep = Left$(varrep, 2) & "/" & Mid$(varrep, 3, 2) & "/" & Right$(varrep, 4)
         'Else
         '   varstr = Right$(varrep, 2)
         '   If Val(varstr) < 30 Then
         '      varstr = "20" & varstr
         '   Else
         '      varstr = "19" & varstr
         '   End If
         '   varrep = Left$(varrep, 2) & "/" & Mid$(varrep, 3, 2) & "/" & varstr
         'End If
      
      Case "E"   'date : DD mois YYYY
         'jld20041102
         varrep = f_temps(varrep, 2)
         If Trim$(varrep) = "" Then varrep = "********"
         If Len(varrep) = 8 Then
            varrep = Left$(varrep, 2) & " " & f_temps(varrep, 10) & " " & Right$(varrep, 4)
         Else
            varstr = Right$(varrep, 2)
            If Val(varstr) < 30 Then
               varstr = "20" & varstr
            Else
               varstr = "19" & varstr
            End If
            varrep = Left$(varrep, 4) & varstr
            varrep = Left$(varrep, 2) & " " & f_temps(varrep, 10) & " " & Right$(varrep, 4)
         End If
      
      'jld20030604
      Case "F"   'date : Jour DD mois YYYY
         
         'jld20041102
         varrep = f_temps(varrep, 2)
         
         If Trim$(varrep) = "" Then varrep = "********"
         
         If Len(varrep) = 8 Then
            varrep = f_temps(varrep, 7) & " " & Left$(varrep, 2) & " " & f_temps(varrep, 10) & " " & Right$(varrep, 4)
         Else
            varstr = Right$(varrep, 2)
            If Val(varstr) < 30 Then
               varstr = "20" & varstr
            Else
               varstr = "19" & varstr
            End If
            varrep = Left$(varrep, 4) & varstr
            varrep = f_temps(varrep, 7) & " " & Left$(varrep, 2) & " " & f_temps(varrep, 10) & " " & Right$(varrep, 4)
         End If
      
      'jld20030604
      Case "G"   'date : nom du Jour
         'jld20041102
         varrep = f_temps(varrep, 2)
         
         If Trim$(varrep) = "" Then varrep = "********"
         
         If Len(varrep) = 8 Then
            varrep = f_temps(varrep, 7)
         Else
            varstr = Right$(varrep, 2)
            If Val(varstr) < 30 Then
               varstr = "20" & varstr
            Else
               varstr = "19" & varstr
            End If
            varrep = Left$(varrep, 4) & varstr
            varrep = f_temps(varrep, 7)
         End If
      
      Case "H"   'heure HH"H"MM 12H34
         If Trim$(varrep) = "" Then varrep = "*****"
         Select Case Len(Trim$(varrep))
         Case 8 'HHxMMxSS"
            varrep = Trim$(f_champ(varrep, ":", 1)) & "H" & Trim$(f_champ(varrep, ":", 2))
         Case 5 'HHxMM
            varrep = Left$(varrep, 2) & "H" & Right$(varrep, 2)
         Case 4 'HHMM
            varvar = varrep
            If IsNumeric(varvar) Then
               varrep = Left$(varrep, 2) & "H" & Right$(varrep, 2)
            End If
         Case Else
         End Select
      
      Case "JJ"  'jour sur 2 caractères
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = Left$(varrep, 2)
      
      Case "L"  'multiligne
      'MsgBox varrep
         varrep = f_remplace(varrep, "||", CrLf & "\par ", "T", 1)

      Case "M"  'numérique avec ___ si rien
         'on remplace le séparateur décimal par le paramètre régional
         If Val(varrep) = 0 Or varrep = "" Then
            varrep = "___"
         Else
            varrep = f_remplace(varrep, vardec, VarDecWin, "TOUS", 1)
         End If

      Case "MM" 'mois sur 2 caractères
         If Trim$(varrep) = "" Then varrep = "********"
         varrep = Mid$(varrep, 3, 2)

      'jld20030604
      Case "MO" 'mois libelle
         If Trim$(varrep) = "" Then varrep = "********"
         
         If Len(varrep) = 8 Then
            varrep = f_temps(varrep, 10)
         Else
            varstr = Right$(varrep, 2)
            If Val(varstr) < 30 Then
               varstr = "20" & varstr
            Else
               varstr = "19" & varstr
            End If
            varrep = Left$(varrep, 4) & varstr
            varrep = f_temps(varrep, 10)
         End If

      Case "N"  'numérique
         If Trim$(varrep) = "" Then varrep = "0"
         'on remplace le séparateur décimal par le paramètre régional
         varrep = f_remplace(varrep, vardec, VarDecWin, "TOUS", 1)
      
      Case "P"  'prénom : première lettre en majuscule Jean-Louis
         varrep = LCase$(varrep)
         varint = f_remplit_str(varrep, " ")
         For i = 1 To varint
            If i = 1 Then
               varrep = UCase$(Left$(tabstr(i), 1)) & Mid$(tabstr(i), 2)
            Else
               varrep = varrep & " " & UCase$(Left$(tabstr(i), 1)) & Mid$(tabstr(i), 2)
            End If
         Next i

         varint = f_remplit_str(varrep, "-")
         For i = 1 To varint
            If i = 1 Then
               varrep = UCase$(Left$(tabstr(i), 1)) & Mid$(tabstr(i), 2)
            Else
               varrep = varrep & "-" & UCase$(Left$(tabstr(i), 1)) & Mid$(tabstr(i), 2)
            End If
         Next i

      Case "R"  'alpha remplacement laprtfrp.txt
         Err = 0
         On Error Resume Next
         nf = FreeFile
         Close #nf: Open lapdic & "laprtfrp.txt" For Input As #nf
         If Err Then
            On Error Resume Next
            Close #nf: Open varlaprap & "laprtfrp.txt" For Input As #nf
         End If
         If Err Then
            MsgBox (Str(Err) & " ERREUR LEC : " & lapdic & "laprtfrp.txt")
         Else
            Do While Not EOF(nf)
               Line Input #nf, a$
               If Trim$(varrep) = "" Then
                  If InStr(1, a$, varsep) <> 0 And Trim$(f_champ(a$, varsep, 1)) = varrep Then
                     varrep = f_champ(a$, varsep, 2)
                     Exit Do
                  End If
               Else
                  If InStr(1, a$, varsep) <> 0 And Trim$(UCase$(f_champ(a$, varsep, 1))) = Trim$(UCase$(varrep)) Then
                     varrep = f_champ(a$, varsep, 2)
                     Exit Do
                  End If
               End If
            Loop
         End If
         Close #nf

      Case "T"   'heure HH:MM
         If Trim$(varrep) = "" Then varrep = "*****"
         Select Case Len(varrep)
         Case 8
            varrep = Trim$(f_champ(varrep, ":", 1)) & ":" & Trim$(f_champ(varrep, ":", 2))
         Case 5
            varrep = Left$(varrep, 2) & ":" & Right$(varrep, 2)
         Case 4
            varrep = Left$(varrep, 2) & ":" & Right$(varrep, 2)
         Case Else
         End Select
      
      'jld20041115 conversion 1/4 1/2 3/4
      Case "LTOD"
         varrep = f_remplace(varrep, "1/4", "0.25", "T", 1)
         varrep = f_remplace(varrep, "1/2", "0.50", "T", 1)
         varrep = f_remplace(varrep, "3/4", "0.75", "T", 1)
         varrep = f_remplace(varrep, "1/3", "0.33", "T", 1)
         varrep = f_remplace(varrep, "2/3", "0.66", "T", 1)

      Case Else
         varrep = varrep

      End Select

      'jld20030710
      varint = 0
      If vartaille <> "" Then
         Select Case vardg
         Case "D"
            varint = Val(vartaille)
            If varint <> 0 Then
               varrep = f_r_d(varrep, varcar, varint)
            End If
         Case "G"
            varint = Val(vartaille)
            If varint <> 0 Then
               varrep = f_r_g(varrep, varcar, varint)
            End If
         'jld20040207 partie de champ
         Case "P"
            varint = Val(vartaille)
            If varint <> 0 Or Trim$(vartaille) = "*" Then
               If Val(varcar) > 0 Then
                  If Trim$(vartaille) = "*" Then
                     varrep = Mid$(varrep, Val(varcar))
                  Else
                     varrep = Mid$(varrep, Val(varcar), varint)
                  End If
               End If
            End If
         'jld20041102 sous valeur sep:numchp:C
         Case "C" 'champ
            varint = Val(vartaille)
            If varint <> 0 Then
               If Trim$(varcar) <> "" Then
                  varrep = f_champ(varrep, varcar, varint)
               End If
            End If
         Case Else
         End Select
      End If


      f_typchp = varrep

End Function

Function f_trtchp(varstr1 As String) As String
'traitement des formules de construction de données à partir de constantes et champs fichier
'varstr1 = ligne de formules champ
'ex:"IDE"+[@CONSEXT;56]
'retour = traduction des formules

   Dim varstr As String
   Dim varrep As String
   Dim varlig As String
   Dim varret As String
   Dim varcar As String
   Dim varint As Integer
   Dim varnbr As Integer
   Dim i As Integer
   Dim j As Integer
   Dim tabtrtchp() As String
   ReDim tabtrtchp(0)
   
   varlig = varstr1
   varret = ""
   varint = 0
   varnbr = 0
   
   varint = f_remplit_str(varlig, "+")
   varnbr = varint
   
   For i = 1 To varnbr
      ReDim Preserve tabtrtchp(i)
      tabtrtchp(i) = tabstr(i)
   Next i
   
   varret = ""
   For i = 1 To varnbr
      
      varstr = ""
      varstr = Trim$(tabtrtchp(i))
      varstr = f_var(varstr)
      varret = varret & varstr
      
   Next i
   
   f_trtchp = varret
   
   
End Function

Function f_fic(prog As String, orific As String, laps As Single, msg As String) As String

   'FRED20040203, controle la modification ou la creation d'un fichier
   'dans un laps de temps donne.
   '
   'prog : programme modifiant le fichier. Si "" alors pas de programme a gerer
   '       Le cas échéant le programme est lancé en mode caché
   'orific : adresse et nom du fichier a controler
   'laps : laps du temps de controle
   'msg : message si le fichier n'a pas ete créé/modifié pendant le laps de temps
   '      defini. Si "" alors pas de message
   '

   Dim datficori As String
   Dim datficdes As String
   Dim t0 As Single
   Dim t1 As Single
   Dim l As String
   Dim x As Integer
      
   On Error Resume Next
      
   f_fic = "OK"
   
   Screen.MousePointer = 11
   DoEvents
   
   datficori = FileDateTime(orific)
   If prog <> "" Then x = Shell(prog, 0)
   Err = 0
   t0 = Timer
   'On attend la realisation du fichier
   If datficori <> "" Then
      datficdes = FileDateTime(orific)
      Do While datficori >= datficdes
         t1 = Timer
         If (Timer Mod t1) = 1 Then
            datficdes = FileDateTime(orific)
         End If
         If t1 >= t0 + laps Then Exit Do
      Loop
      If t1 >= t0 + laps Then
         If msg <> "" Then MsgBox msg
         f_fic = ""
         GoTo FIN
      End If
   Else
      l = Dir$(orific)
      Do While l = ""
         t1 = Timer
         If (Timer Mod t1) = 1 Then
            l = Dir$(orific)
         End If
         If t1 >= t0 + 7 Then Exit Do
      Loop
      If t1 >= t0 + 7 Then
         If msg <> "" Then MsgBox msg
         f_fic = ""
         GoTo FIN
      End If
   End If
   
FIN:
   Screen.MousePointer = 1
   DoEvents

End Function


Function f_lecver(varstr1 As String, varstr2 As String) As Integer
'jld20040621 lecture du fichier .ver
'varstr1 = masque
'varstr2 = M/A si M alors masque principal  sinon autres (masque dans règles etc)
'retour = nombre de valeurs trouvées : données dans tableau tabverm ou tabvera

   Dim varche As String
   Dim vartyp As String
   Dim varstr As String
   Dim varres As String
   Dim vardir As String
   Dim varfic As String
   Dim varlig As String
   Dim varmsq As String
   Dim varindche As Integer
   Dim varmillier As Integer
   Dim mapsuffixe As String
   Dim varnbr As Integer
   Dim varint As Integer
   Dim nf As Integer
   Dim i As Integer
   
   f_lecver = 0
   
   vartyp = Left$(Trim$(UCase$(varstr2)), 1)
   
   If vartyp = "M" Then
      ReDim tabverm(0)
      varmmi = lapmmi
      varsuf = lapsuf
   Else
      ReDim tabvera(0)
      mapmmi = lapmmi
      mapsuf = lapsuf
   End If
   
   varmsq = Trim$(UCase$(varstr1))
   If varmsq = "" Then
      Exit Function
   End If
   
   varfic = varmsq & ".VER"
   varlig = ""
   
   On Error Resume Next
   nf = FreeFile
   Close #nf: Open lapdic & varfic For Input As #nf
   If Err Then
      If Err <> 53 And Err <> 75 Then
         MsgBox Str(Err) & " ERREUR LECTURE fichier : " & varfic
         Exit Function
      End If
   Else
      Line Input #nf, varlig
   End If
   Close #nf
   
   varlig = Trim$(varlig)
   If varlig <> "" Then
      varnbr = 0
      varnbr = f_remplit_str(varlig, varsep)
      If varnbr >= 2 Then
         'varstr = ""
         'varindche = 0
         'varstr = Trim$(f_champ(varlig, varsep, 2))
         'varindche = Val(varstr)
         'vardir = ""
         'If varstr <> "" And varindche > 0 And varindche <= 30 Then
         '   vardir = Trim$(di$(varindche))
         '   If Right$(vardir, 1) <> "\" Then
         '      vardir = vardir & "\"
         '   End If
         '   tabstr(2) = vardir
         'End If
      Else
         Exit Function
      End If
      
      Select Case vartyp
      Case "M"
         For i = 1 To varnbr
            ReDim Preserve tabverm(i)
            tabverm(i) = tabstr(i)
            
            Select Case i
            Case 1 'masque
                nommasar = Trim$(tabverm(i))
                If InStr(nommasar, "*") <> 0 Then
                   nomg = Left$(nommasar, InStr(nommasar, "*") - 1)
                   nomd = Right$(nommasar, Len(nommasar) - InStr(nommasar, "*"))
               End If
            
            Case 2 'chemin
                varche = Trim$(tabverm(i))
                indexchemin = Val(varche)
                If indexchemin = 0 And varche <> "" Then
                  If Right$(varche, 1) <> "\" Then
                     varche = varche & "\"
                  End If
                  varche = f_adr(varche)
                  di$(2) = varche
                  indexchemin = 2
                End If
                If indexchemin = 0 Then
                  indexchemin = 1
                End If
                tabverm(i) = indexchemin

            Case 3 'millier
               varstr = ""
               varstr = Trim$(tabverm(i))
               If varstr <> "" Then
                  varmmi = Val(varstr)
               End If

            Case 4 'suffixe
                 varstr = ""
                 varsuf = ".std"
                 varstr = Trim$(tabverm(i))
                 If varstr <> "" Then
                    If Left$(varstr, 1) <> "." Then
                     varstr = "." & varstr
                    End If
                    varsuf = varstr
                 End If
                                         
            Case Else
            End Select
         
         Next i
      
      Case Else
         For i = 1 To varnbr
            ReDim Preserve tabvera(i)
            tabvera(i) = tabstr(i)
         
            Select Case i
            Case 1  'masque
                'chgt nom masque
                mapmsq = Trim$(tabvera(i))
            
            Case 2  'indexchemin
               varche = Trim$(tabvera(i))
               varint = Val(varche)
               If varche <> "" Then
                  If varint = 0 Then
                     If Right$(varche, 1) <> "\" Then
                        varche = varche & "\"
                     End If
                     varche = f_adr(varche)
                     tabvera(i) = varche
                  Else
                     If varint > 0 And varint <= 30 Then
                        tabvera(i) = di$(varint)
                     End If
                  End If
               Else
                  tabvera(i) = di$(1)
               End If
               mapmed = tabvera(i)

            Case 3  'millier
               mapmmi = lapmmi
               varstr = ""
               varstr = Trim$(tabvera(i))
               If varstr <> "" Then
                  mapmmi = Val(varstr)
               End If

            Case 4  'suffixe
                 varstr = ""
                 mapsuf = ".std"
                 varstr = Trim$(tabvera(i))
                 If varstr <> "" Then
                    mapsuf = varstr
                 End If
                                         
            Case Else
            End Select
         
         Next i
      End Select
      
   Else
      Exit Function
   End If

   f_lecver = varnbr
   
End Function
Function f_cmp(varstr1 As String, varstr2 As String, varint1 As Integer, varint2 As Integer) As String
'jld20040628 test deux valeurs : f_cmp("BLANC","!NOIR*",1)
'varstr1 = valeur1
'varstr2 = valeur2
'varint1 = type de comparaison : 0=partie 1=exact
'varint2 = type de comparaison : 1=non case sensitif
'retour = OK si test verifié

   Dim varstr As String
   Dim varinv As Integer
   Dim vareto As Integer
   Dim vartst As String
   Dim varcas As Integer
   Dim varstrg As String
   Dim varstrd As String
   Dim varcar As String
   Dim varinf As Integer
   Dim varlen As Integer
   Dim varchpg As Integer
   Dim varchpd As Integer
   Dim varexc As Integer
   Dim varint As Integer
   Dim varnbr As Integer
   
   f_cmp = ""
   
   varstrd = Trim$(varstr2)
   varexc = varint1
   varcas = varint2
         
   varint = 0
   varnbr = 0
   varinf = 0
   
   varint = f_remplit_str(varstrd, "&")
   For varnbr = 1 To varint
      varstrg = Trim$(varstr1)
      varstrd = Trim$(tabstr(varnbr))
      varinv = 0
      vareto = 0
      varstr = ""
      
      'test caractère !
      varcar = ""
      varcar = Left$(Trim$(varstrd), 1)
      If varcar = "!" Then
         varinv = 1
         varstrd = f_champ(varstrd, "!", 2)
      Else
         varinv = 0
      End If
   
      'test caractère *
      varcar = ""
      varcar = Right$(Trim$(varstrd), 1)
      If varcar = "*" Then
         vareto = 1
         varstrd = Left$(varstrd, Len(varstrd) - 1)
      End If
      
      If vareto = 1 Then
         varstrg = Left$(varstrg, Len(varstrd))
      End If
            
      'jldyyy gestion des blancs à voir
      Select Case varcas
      Case 1
         varstrg = Trim$(UCase$(varstrg))
         varstrd = Trim$(UCase$(varstrd))
      Case Else
         varstrg = Trim$(varstrg)
         varstrd = Trim$(varstrd)
      End Select
      
      'partie de mot
      If varexc = 0 Then
         If varstrg = "" And varstrd = "" Then
            varstrg = "@"
            varstrd = "@"
         End If
         If varstrd <> "" Then
            If InStr(varstrg, varstrd) > 0 Then
               varstrg = varstrd
            End If
         Else
         End If
         
      End If
   
      Select Case varinv
      Case 1
         If varstrg <> varstrd Then
            If varnbr = 1 Then
               varinf = 1
            Else
               If varinf = 1 Then
                  varinf = 1
               End If
            End If
         Else
            varinf = 0
         End If
      Case Else
         If varstrg = varstrd Then
            If varnbr = 1 Then
               varinf = 1
            Else
               If varinf = 1 Then
                  varinf = 1
               End If
            End If
         Else
            varinf = 0
         End If
      End Select
   Next varnbr
   
   If varinf = 1 Then
      f_cmp = "OK"
   End If

   'jld20040816 remise à zéro err si erreur 13 : date avec 00 (jour mois)
   Err = 0
   
End Function
