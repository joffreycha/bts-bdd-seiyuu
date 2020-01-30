Attribute VB_Name = "Commun"
Option Explicit
Public Const TOUT = 1
Public Const MAJUSCULES = 2
Public Const CHIFFRES = 3
Public Const CHIFFRESSPACE = 4
Public Const CHIFFRESMOINS = 5
Public Const MONNAIE = 6
Public Const DATES = 7
Public Const INITIALE = 8
Public Const TELEPHONE = 9
Public Const URL = 10
Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Declare Function DiskSpaceFree Lib "STKIT432.DLL" () As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pPrinter As Any, ByVal cbBuf As Long, pcbNeeded As Long) As Boolean
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, pDefault As Any) As Boolean
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Type POINTAPI
  x As Integer
  y As Integer
End Type

Public Function BeginMonth(madate As Date) As Date
  'V 1.0
  'Paramètres : une date
  'Retour : date du premier jour du mois
  BeginMonth = DateSerial(Year(madate), Month(madate), 1)
End Function

Public Function bin$(ByVal x)
  'V 1.0
  'Paramètres : un nombre positif
  'Retour : chaine représentant le nombre en binaire sur 24 bits
  Dim i%
  For i% = 24 To 1 Step -1
    If x >= 2 ^ (i% - 1) Then
      bin$ = bin$ & "1"
      x = x - 2 ^ (i% - 1)
    Else
      bin$ = bin$ & "0"
    End If
  Next
End Function

Function BinToDéc(bin$)
  'V 1.0
  'Paramètres : Chaine représentant un nombre binaire de 31 bits maximum
  'Retour : un entier long
  Dim i%
  For i% = 1 To Len(bin$)
    If Mid$(bin$, i%, 1) = "1" Then BinToDéc& = BinToDéc& + 2 ^ Abs(i% - Len(bin$))
  Next i%
End Function

Public Function BitMax%(x%)
  'V 1.1
  'Paramètre : un entier positif
  'Action : Le N° (0 à 15) du bit de poids le plus fort qui est à 1
  Dim i%
  For i% = 15 To 0 Step -1
    If x% >= 2 ^ i% Then Exit For
  Next
  Bit_Max% = i%
End Function

Public Sub CentreFocus(focus As Control, picture As Control)
  'V 1.0
  'Paramètres : le controle à centrer
  '             le controle de référence
  'Action : Le premier controle est centré sur le second horizontalement et verticalement.
  focus.Left = picture.Left - ((focus.Width - picture.Width) / 2)
  focus.Top = picture.Top - ((focus.Height - picture.Height) / 2)
End Sub

Public Function chemin$(chaine$)
  'V 1.0
  'Paramètres : Un nom de fichier et son chemin
  'Retour : Le nom du chemin seul terminé par \
  Dim dummy$, x%
  dummy$ = chaine$
  chemin$ = ""
  x% = InStr(dummy$, "\")
  Do While x% > 0
    x% = InStr(x% + 1, dummy$, "\")
    If x% > 0 Then chemin$ = Left$(dummy$, x%)
  Loop
End Function

Public Function ComputerName$()
  'V 2.1
  'Retour : Le nom d'ordinateur dans le réseau
  Dim retour$
  retour$ = Space$(50)                           'DOIT être initialisé
  If GetComputerName(retour$, 50) <> 0 Then ComputerName$ = Mid$(retour$, 1, InStr(retour$, Chr$(0)) - 1) Else ComputerName$ = ""
End Function

Function DécToHex$(ByVal x)
  'V 1.0
  'Paramètres : un nombre
  'Retour : une chaine représentant le nombre en hexadécimal
  DécToHex$ = Val("&H" & Str$(x))
End Function

Public Function EmptyToZéro$(valeur$)
  'V 1.1
  'Paramètre : une chaine représentant un nombre ou une chaine vide
  'Retour : La chaine ou "O"
  If x$ = "" Then EmptyToZéro$ = 0 Else EmptyToZéro$ = CDbl(x$) 'Peux pas utiliser IIF qui évalue tout
End Function

Public Function EndMonth(madate As Date) As Date
  'V 1.0
  'Paramètres : une date
  'Retour : date du dernier jour du mois
  EndMonth = DateSerial(Year(madate), Month(madate) + 1, 1) - 1
End Function

Public Function Exist(nom_fichier$) As Boolean
  'V 1.1
  'Paramètres : un nom de fichier
  'Retour : Vrai si le fichier existe, sinon Faux
 Exist = IIf(Len(Dir(nom_fichier$)) > 0, True, False)
End Function

Public Sub FiltreSaisie(z As TextBox, x%, mode%)
  'V 3.1
  'Paramètres : un controle TEXTBOX
  '             le code% ASCII d'un caractère
  '             un entier
  'Action : Le texte de TEXTBOX est filtré selon différents critères
  'Mode% 1 Tout car., pas d'espace en 1ere position
  'Mode% 2 Transforme en majuscules
  'Mode% 3 Chiffres seuls
  'Mode% 4 Chiffres et espace
  'Mode% 5 Chiffres et -
  'Mode% 6 Valeur monétaire
  'Mode% 7 Date
  'Mode% 8 Initiale en majuscule
  'Mode% 9 Numéro de téléphone
  'Mode% 10 Url
  Dim i%
  Dim last As Boolean
  last = IIf(z.SelStart <= Len(z.Text) - 1, False, True) 'Vrai si le point d'insertion est derrière le dernier caractère
  If z.SelStart = 0 And x% = 32 Then                    'Elimine espace en 1ère position
    x% = 0
    Exit Sub
  End If
  If x% = 8 Then Exit Sub                               'Backspace toujours actif
  Select Case mode%
  Case TOUT
    'Sauf espace en 1ère position
  Case MAJUSCULES
    x% = Asc(UCase$(Chr(x)))
  Case INITIALE
    If z.SelStart = 0 Then x% = Asc(UCase$(Chr$(x%)))
  Case CHIFFRES
    If NotNumber(x%) Then x% = 0
  Case CHIFFRESSPACE
    If NotNumber(x%) And x% <> 32 Then x% = 0
  Case CHIFFRESMOINS
    If NotNumber(x%) And x% <> 45 Then x% = 0
  Case MONNAIE
    If x% = 46 Then x% = 44 'Transforme point en virgule
    If NotNumber(x%) And x% <> 44 And x% <> 45 Then x% = 0 'Ne garder que les chiffres, la virgule et le '-'
    If z.SelLength = Len(z.Text) Then Exit Sub
    'Empêcher le '-' ailleurs qu'en première colonne
    If x% = 45 And z.SelStart > 0 Then x% = 0
    'Empêcher plus d'un '-'
    If x% = 45 And InStr(z.Text, "-") > 0 Then x% = 0
    'Empêcher une frappe avant le '-'
    If z.SelStart = 0 And InStr(z.Text, "-") > 0 Then x% = 0
    'Empêcher plus d'une virgule
    If x% = 44 And InStr(z.Text, ",") > 0 Then x% = 0
    'Empêcher virgule décimale avant plus de 2 décimales
    i% = InStr(z.Text, " ") - 1 'Pour éliminer le symbole monétaire qui suit l'espace
    If i% = -1 Then i% = Len(z.Text) Else i% = i%
    If x% = 44 And z.SelStart < i% - 2 Then x% = 0
    'Empêcher plus de deux décimales derrière la virgule
    i% = InStr(z.Text, ",")
    If Not NotNumber(x%) And i% > 0 And z.SelStart >= i% And Len(z.Text) >= i% + 2 Then x% = 0
  Case DATES
    If NotNumber(x%) Then x% = 47 'Transforme tout autre car. qu'un chiffre en '/'
    If z.SelLength = Len(z.Text) Then Exit Sub
    'Empêcher plus de 2 '/'
    i% = InStr(z.Text, "/")
    If x% = 47 And i% > 0 Then If InStr(i% + 1, z.Text, "/") > 0 Then x% = 0
    'Empêcher un '/' en colonne 1 ou > à 6
    If x% = 47 And (z.SelStart = 0 Or z.SelStart > 5) Then x% = 0
    'Empêcher 2 '/' consécutifs
    If x% = 47 And z.SelStart > 0 Then If Asc(Mid$(z.Text, z.SelStart)) = 47 Then x% = 0 'Regarde à gauche
    If x% = 47 And Not last Then If Asc(Mid$(z.Text, z.SelStart + 1)) = 47 Then x% = 0 'Regarde à droite
    'Empêcher plus de 2 chiffres pour le jour
    If x% <> 47 And z.SelStart <= 2 And Len(z.Text) >= 2 And IsNumeric(Left$(z.Text, 2)) Then x% = 0
    'Empêcher plus de 2 chiffres pour le mois
    i% = InStr(z.Text, "/")
    If x% <> 47 And i% > 0 And z.SelStart >= i% And z.SelStart <= i% + 2 And Len(z.Text) >= i% + 2 And IsNumeric(Mid$(z.Text, i% + 1, 2)) Then x% = 0
    'Empêcher plus de 4 chiffres pour l'année
    i% = InStr(z.Text, "/")
    If x% <> 47 And i% > 0 Then
      i% = InStr(i% + 1, z.Text, "/")
      If i% > 0 And z.SelStart >= i% And Len(z.Text) >= i% + 4 Then x% = 0
    End If
  Case TELEPHONE
    If x% = 46 Then x% = 32 'Transforme point en espace
    If NotNumber(x%) And x% <> 43 And x% <> 32 Then x% = 0  'Ne garder que les chiffres, le + et l'espace
    If z.SelLength = Len(z.Text) Then Exit Sub
    'Empêcher le '+' ailleurs qu'en première colonne
    If x% = 43 And z.SelStart > 0 Then x% = 0
    'Empêcher plus d'un '+'
    If x% = 43 And InStr(z.Text, "+") > 0 Then x% = 0
    'Empêcher une frappe avant le '+'
    If z.SelStart = 0 And InStr(z.Text, "+") > 0 Then x% = 0
    'Empêcher 2 espaces consécutifs
    If x% = 32 And z.SelStart > 0 Then If Asc(Mid$(z.Text, z.SelStart)) = 32 Then x% = 0 'Regarde à gauche
    If x% = 32 And Not last Then If Asc(Mid$(z.Text, z.SelStart + 1)) = 32 Then x% = 0 'Regarde à droite
    'Empêcher un espace après le '+'
    If x% = 32 And z.SelStart > 0 Then If Asc(Mid$(z.Text, z.SelStart)) = 43 Then x% = 0
  Case URL
    If Len(z.Text) >= 4 Then
      If Left$(z.Text, 4) = "http" Then
        MsgBox "Ne tapez pas : 'http://'", vbOKOnly + vbExclamation, ""
        x% = 0
        z.Text = IIf(Len(z.Text) = 4, "", Mid$(z.Text, 5))
      End If
    End If
    Select Case x%
    Case 45 To 57
    Case 65 To 90
    Case 97 To 122
    Case 232, 233
      x% = 101
    Case 224
      x% = 97
    Case 231
      x% = 99
    Case Else
      x% = 0
    End Select
  End Select
End Sub

Public Function FormatTél$(dummy$)
  'V 1.0
  'Paramètre : une chaine représentant un numéro de téléphone
  'Retour : La chaine dont le 00 de début est remplacée par un '+'
  FormatTél$ = dummy$
  If Len(dummy$) > 4 Then   'Moins de 4 chiffres n'a pas de sens
    If Left$(dummy$, 2) = "00" Then FormatTél$ = "+" & Mid$(dummy$, 3)
    If Left$(dummy$, 3) = "00 " Or Left$(dummy$, 3) = "0 0" Then FormatTél$ = "+" & Mid$(dummy$, 4)
    If Left$(dummy$, 4) = "0 0 " Then FormatTél$ = "+" & Mid$(dummy$, 5)
  End If
End Function

Function GetDiskSpaceFree(ByVal strDrive As String) As Long
  'V 1,0
  'Paramètres : Un nom de lecteur
  'Retour : La quantité d'espace libre, ou -1 si une erreur s'est produite.
  'DLL nécessaire : STKIT416.DLL
  Dim strCurDrive As String
  Dim lDiskFree As Long
  On Error Resume Next
  ' Enregistre le lecteur en cours.
  strCurDrive = Left$(CurDir$, 2)
  strDrive = strDrive & ":"
  ' Change le lecteur par défaut. La fonction API DiskSpaceFree() utilise uniquement le lecteur par défaut.
  ChDrive strDrive
  ' S'il n'est pas possible de changer de lecteur par  défaut, c'est une erreur.
  If Err.Number <> 0 Or (strDrive <> Left$(CurDir$, 2)) Then
  lDiskFree = -1
  Else
  lDiskFree = DiskSpaceFree()
  If Err.Number <> 0 Then lDiskFree = -1 ' Si DiskSpaceFree provoque une erreur.
  End If
  GetDiskSpaceFree = lDiskFree
  ' Remplace le lecteur en cours.
  ChDrive strCurDrive
  Err.Number = 0
End Function

Function HexToDéc&(ByVal hexa$)
  'V 1.0
  'Paramètres : une chaine représentant le nombre en hexadécimal
  'Retour : le nombre en décimal
  HexToDéc& = Val("&H" & hexa$)
End Function

Public Function IsInteger(nombre) As Boolean
  'V 1.0
  'Paramètre : un nombre
  'Retour : Vrai si le nombre est entier
  IsInteger = IIf(nombre - Int(nombre) = 0, True, False)
End Function

Public Function Max(a, b)
  'V 1.1
  'Paramètres : deux nombres
  'Retour : Le plus grand des deux nombre
  Max = IIf(a > b, a, b)
End Function

Public Sub MidPrint(obj As Object, texte$, largeur%)
  'V 1.0
  'Paramètre : un objet recevant du texte, le texte, la largeur de colonne
  ''Action : Le texte est écrit au milieu de la colonne
  obj.CurrentX = obj.CurrentX + (largeur% - obj.TextWidth(texte$)) / 2
  obj.Print texte$;
End Sub

Public Function Min(a, b)
  'V 1.1
  'Paramètres : deux nombres
  'Retour : Le plus petit des deux nombre
 Min = IIf(a < b, a, b)
End Function

Public Function NomFichier$(chemin$)
  'V 1.1
  'Paramètres : Un nom de fichier et son chemin
  'Retour : Le nom du fichier sans le chemin
  Dim dummy$, x%
  dummy$ = chemin$
  x% = InStr(dummy$, "\")
  Do While x% > 0
    dummy$ = Mid$(dummy$, x% + 1)
    x% = InStr(dummy$, "\")
  Loop
  NomFichier$ = dummy$
End Function

Public Function NoNul$(valeur)
  'V 1.1
  'Paramètre : une chaine ou NULL
  'Retour : La chaine ou une chaine vide si NULL en entrée
  If IsNull(valeur) Then NoNul = "" Else NoNul = valeur
End Function

Public Function NotNumber(code%) As Boolean
  'V 1.0
  'Paramètre : un code ASCII
  'Retour : Vrai si ce n'est pas le code d'un chiffre (0 -9)
  NotNumber = IIf(code% < 48 Or code% > 57, True, False)
End Function

Public Function NoVide$(valeur)
  'V 1.1
  'Paramètre : une chaine représentant un nombre ou une chaine vide
  'Retour : La chaine ou "O"
  If Len(valeur) = 0 Then NoVide = "0" Else NoVide = valeur
End Function

Public Function Nozéro$(valeur$)
  'V 1.1
  'Paramètre : une chaine représentant un nombre
  'Retour : La chaine ou chaine vide si nombre = O
  If CDbl(valeur$) = 0 Then Nozéro = "" Else Nozéro = valeur$
End Function

Public Function NullToString$(valeur)
  'V 1.0
  'Paramètre : une chaine ou NULL
  'Retour : La chaine ou une chaine vide si NULL en entrée
  NullToString$ = IIf(IsNull(valeur), "", valeur)
End Function

Public Sub Position(fen As Form, dial_box As Form)
  'V 2.0
  'Paramètres : la fenêtre MDI ou une fenêtre fille ou une fenêtre modale
  '             une boite de dialogue (fenêtre modale)
  'Action : La boite de dialogue est centrée dans la fenêtre.
  Dim margehaute%, margegauche%, coord As POINTAPI, décal%, con As Control
  If fen.MDIChild = True Then          'fen = fenêtre fille
    coord.x = 0
    coord.y = 0
    Call ClientToScreen(Forms(0).hwnd, coord)
    On Error Resume Next               'Tous les controles n'ont pas une propriété ALIGN
    For Each con In Forms(0).Controls  'On cherche le décalage dû à des barres d'outils
      If con.Align = vbAlignTop Then
        If Err = 0 Then décal% = décal% + con.Height
      End If
    Next
    On Error GoTo 0
    margehaute% = coord.y * Screen.TwipsPerPixelY + décal%
    margegauche% = coord.x * Screen.TwipsPerPixelX
  End If
  dial_box.Top = (fen.Height - dial_box.Height) / 2 + fen.Top + margehaute%
  dial_box.Left = (fen.Width - dial_box.Width) / 2 + fen.Left + margegauche%
  If dial_box.Top < 0 Then dial_box.Top = 0
  If dial_box.Left < 0 Then dial_box.Left = 0
  If dial_box.Top > Screen.Height - dial_box.Height Then dial_box.Top = Screen.Height - dial_box.Height
  If dial_box.Left > Screen.Width - dial_box.Width Then dial_box.Left = Screen.Width - dial_box.Width
End Sub

Public Sub PrintLeftWithDot(obj As Object, texte$, largeur%, motif$, largeur_motif%)
  'V 1.0
  'Paramètre : un objet recevant du texte, le texte, la largeur de colonne, une chaine 'motif' et la largeur de cette chaine
  'Action : Le texte est écrit à gauche et le motif est répété sur toute la largeur de la colonne
  Dim i%, x%
  texte$ = RTrim$(texte$)
  x% = Int(largeur% - obj.TextWidth(texte$) / largeur_motif%)
  For i% = 0 To x%
    texte$ = texte$ & motif$
  Next
  obj.Print texte$;
End Sub

Public Sub RightPrint(obj As Object, texte$, largeur%)
  'V 1,0
  'Paramètre : un objet recevant du texte, le texte, la largeur de colonne
  'Action : Le texte est écrit justifié à droite de la colonne
  obj.CurrentX = obj.CurrentX + (largeur% - obj.TextWidth(texte$))
  obj.Print texte$;
End Sub

Public Sub SelectText(con As TextBox)
  'V 2.0
  'Paramètre : un control TextBox
  'Action : Le texte du control est sélectionné
  con.SelStart = 0
  con.SelLength = Len(con.Text)
End Sub

Public Function ShareName$(nom$)
  'V 3.0
  'Paramètre : Une chaine représentant le DeviceName
  'Retour : Le nom de partage
  'Requis : OpenPrinter, GetPrinter, ClosePrinter
  Dim hPrinter&, BufferLen&, requis&, buffer() As Byte, ptr&, dummy$
  ShareName$ = ""
  Call OpenPrinter(nom$, hPrinter&, ByVal 0)
  Call GetPrinter(hPrinter&, 2, 0, 0, requis&)
  If requis& Then
    BufferLen& = requis&
    requis& = 0
    ReDim buffer(BufferLen&)
    Call GetPrinter(hPrinter&, 2, buffer(0), BufferLen&, requis&)
    'buffer contient des pointeurs sur 4 octets , on récupère le 3ème pointeur
    ptr& = buffer(8) + 256# * buffer(9) + 65536 * buffer(10) + 16777216 * buffer(11)
    dummy$ = Space$(32)                          'DOIT être initialisé
    Call CopyMemory(ByVal dummy$, ByVal ptr&, 32)
    ptr& = InStr(dummy$, Chr$(0))
    If ptr& > 0 Then ShareName$ = Left$(dummy$, ptr& - 1)
  End If
  Call ClosePrinter(hPrinter&)
End Function

Public Sub Sort(tableau, i, d, début, fin)
  'V 1.0
  'Paramètres : un tableau
  '             4 entiers
  'Action : Le tableau est trié sur l'index I de la dimension D (D=1 ou 2)
  'début et fin indique la variation de l'index de lignes
  Dim sorted%, j%, k%, swap
  sorted% = False
  Do While Not sorted
    sorted = True
    If d = 2 Then
      For j% = début To fin - 1
        If tableau(j%, i) > tableau(j% + 1, i) Then
          For k% = LBound(tableau, 2) To UBound(tableau, 2)
            swap = tableau(j%, k%)
            tableau(j%, k%) = tableau(j% + 1, k%)
            tableau(j% + 1, k%) = swap
          Next
          sorted = False
        End If
      Next
    Else
      For j% = début To fin - 1
        If tableau(i, j%) > tableau(i, j% + 1) Then
          For k% = LBound(tableau, 1) To UBound(tableau, 1)
            swap = tableau(k%, j%)
            tableau(k%, j%) = tableau(k%, j% + 1)
            tableau(k%, j% + 1) = swap
          Next
          sorted = False
        End If
      Next
    End If
  Loop
End Sub

Public Function TestCleRIB(rib$) As Boolean
  'V 1,0
  'Paramètre : une chaine de 23 caractères représentant un RIB (Lettres en capitales)
  'Retour : Vrai si la clé du RIB est correcte, faux sinon.
  Dim i%, car$
  For i% = 0 To 23              'Remplacement lettres par chiffres
    car$ = Mid$(rib$, i% + 1, 1)
    Select Case car$
      Case "A", "J"
        Mid(rib$, i% + 1, 1) = "1"
      Case "B", "K", "S"
        Mid(rib$, i% + 1, 1) = "2"
      Case "C", "L", "T"
        Mid(rib$, i% + 1, 1) = "3"
      Case "D", "M", "U"
        Mid(rib$, i% + 1, 1) = "4"
      Case "E", "N", "V"
        Mid(rib$, i% + 1, 1) = "5"
      Case "F", "O", "W"
        Mid(rib$, i% + 1, 1) = "6"
      Case "G", "P", "X"
        Mid(rib$, i% + 1, 1) = "7"
      Case "H", "Q", "Y"
        Mid(rib$, i% + 1, 1) = "8"
      Case "I", "R", "Z"
        Mid(rib$, i% + 1, 1) = "9"
    End Select
  Next
  TestCleRIB = IIf(97 - (62 * CLng(Mid$(rib$, 1, 7)) + 34 * CLng(Mid$(rib$, 8, 7)) + 3 * CLng(Mid$(rib$, 15, 7))) Mod 97 = CInt(Mid$(rib$, 22, 2)), True, False)
End Function

Public Function UserName$()
  'V 1.1
  'Retour : Le nom d'utilisateur de la session
  Dim retour$
  retour$ = Space$(50)                           'DOIT être initialisé
  If GetUserName(retour$, 50) <> 0 Then UserName$ = Mid$(retour$, 1, InStr(retour$, Chr$(0)) - 1) Else UserName$ = ""
End Function

Public Function ValideCB(chaine$) As Boolean
  'V 1,0
  'Paramètre : une chaine contenant les 16 chiffres d'un numéro de C.B.
  'Retour : Vrai si le numéro est cohérent
  Dim somme%, i%, digit%
  For i% = 1 To 16 Step 2
    digit% = Mid$(chaine$, i%, 1)
    somme% = somme% + IIf(digit% >= 5, digit% * 2 + 1, digit% * 2)
    somme% = somme% + Mid$(chaine$, i% + 1, 1)
  Next
  ValideCB = somme% Mod 10 = 0
End Function

Public Function ValideDate(x As TextBox, monformat$) As Boolean
  'V 2.0
  'Paramètre : un controle text
  'Retour : Faux si le controle ne contient pas une date, Vrai sinon et la date est reformatée
  If IsNumeric(Right$(x.Text, 4)) And IsDate(x.Text) Then
    x.Text = Format(CDate(x.Text), monformat$)
    ValideDate = True
  Else
    ValideDate = False
  End If
End Function

Public Sub ValideItem(cont As ListBox, source As TextBox)
  'V 1.2
  'Paramètres : un controle ListBox
  '             un controle TextBox
  'Action : Le texte est cherché dans la liste et l'index de liste est positionné.
  'Au départ l'index de liste DOIT être égal à -1.
  Dim i%
  If cont.ListIndex = -1 Then
    i% = 0
    Do While i% < cont.ListCount
      If cont.List(i%) = source.Text Then
        cont.ListIndex = i%
        Exit Do
      End If
      i% = i% + 1
    Loop
  End If
End Sub

Public Function ValideMontant(x As TextBox) As Boolean
  'V 1,1
  'Paramètre : un controle TextBox
  'Retour : Faux si le controle ne contient pas un montant, Vrai sinon et le montant est reformaté
  If x.Text = "" Then x.Text = "0"
  If IsNumeric(x.Text) Then
    x.Text = Format(CCur(x.Text), "currency")
    ValideMontant = True
  Else
    ValideMontant = False
  End If
End Function

Public Function ValideURL(x As TextBox) As Boolean
  'V 1.0
  'Paramètre : un controle text
  'Retour : Faux si le controle ne contient pas un URL, Vrai sinon
  Dim i%
  ValideURL = True
  i% = Len(x)
  Do While i% > 0
    Select Case Asc(Mid$(x, i%, 1))
      Case 45 To 57
      Case 65 To 90
      Case 97 To 122
      Case Else
        ValideURL = False
    End Select
    i% = i% - 1
  Loop
End Function

Public Function VideToZero(x$)
  'V 1.1
  'Paramètres : une chaine representant une valeur ou vide
  'Retour : La valeur ou 0
  If x$ = "" Then VideToZero = 0 Else VideToZero = CDbl(x$) 'Peux pas utiliser IIF qui évalue tout
End Function

Public Function ZéroToEmpty$(valeur$)
  'V 1.0
  'Paramètre : une chaine représentant un nombre
  'Retour : La chaine ou chaine vide si nombre = O
  ZéroToEmpty$ = IIf(CDbl(valeur$) = 0, "", valeur$)
End Function
