Attribute VB_Name = "Mod_Comptabilite"
' ============================================================
'  MODULE : Mod_Comptabilite
'  Projet : Système de gestion comptable SEB
'  Usage   : Importer via Alt+F11 > Fichier > Importer
'            puis exécuter Setup_Boutons (Alt+F8)
' ============================================================

Option Explicit

' -- Constantes ---------------------------------------------------------------
Private Const SCRIPT_PDF    As String = "pdf_extractor.py"
Private Const JSON_TEMP     As String = "extraction_temp.json"
Private Const SH_FOURN      As String = "Fournisseurs_DB"
Private Const SH_CLIENTS    As String = "Clients_DB"
Private Const SH_PDF        As String = "PDF_Import"
Private Const SH_TVA        As String = "TVA_Mensuelle"
Private Const SH_PIVOT      As String = "Pivot_Analytics"

' Couleurs
Private Const CLR_BLEU_FONCE As Long = 5013504    ' #1F4E79 -> RGB(31,78,121)
Private Const CLR_BLEU_MOY   As Long = 11703214   ' #2E75B6 -> RGB(46,117,182)
Private Const CLR_BLANC      As Long = 16777215
Private Const CLR_VERT       As Long = 3624755    ' #375623
Private Const CLR_ROUGE      As Long = 12517376   ' #C00000
Private Const CLR_ORANGE     As Long = 12905233   ' #C55A11
Private Const CLR_VIOLET     As Long = 7340032    ' #7030A0


' ============================================================
'  SECTION 1 — SETUP INITIAL
' ============================================================

Public Sub Setup_Boutons()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Call _Creer_Boutons_PDF_Import
    Call _Creer_Bouton_DB(SH_FOURN,   "Ouvrir_PDF_Ligne",  "Ouvrir PDF",       "D1", CLR_BLEU_FONCE, 120, 26)
    Call _Creer_Bouton_DB(SH_CLIENTS, "Ouvrir_PDF_Ligne",  "Ouvrir PDF",       "D1", CLR_VERT,        120, 26)
    Call _Creer_Bouton_Pivot
    Call _Creer_Bouton_TVA

    Application.ScreenUpdating = True
    MsgBox "Configuration terminée !" & vbCrLf & _
           "Les boutons ont été créés sur toutes les feuilles.", _
           vbInformation, "Setup réussi"
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Erreur lors du setup : " & Err.Description, vbCritical, "Erreur"
End Sub

' -- Boutons feuille PDF_Import ------------------------------------------------

Private Sub _Creer_Boutons_PDF_Import()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PDF)

    ' Supprimer anciens boutons
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ' Position de référence
    Dim topD5  As Double : topD5  = ws.Range("D5").Top
    Dim leftD5 As Double : leftD5 = ws.Range("D5").Left
    Dim top19  As Double : top19  = ws.Range("B19").Top
    Dim leftB  As Double : leftB  = ws.Range("B19").Left
    Dim leftC  As Double : leftC  = ws.Range("C19").Left

    ' Bouton Importer PDF
    _AjouterBouton ws, "btn_ImporterPDF", "Importer PDF", _
        leftD5, topD5, 140, 26, CLR_BLEU_FONCE, "Mod_Comptabilite.ImporterPDF"

    ' Bouton Extraire données
    _AjouterBouton ws, "btn_Extraire", "Extraire donnees", _
        leftB, top19, 145, 34, CLR_BLEU_MOY, "Mod_Comptabilite.ExtraireViaPython"

    ' Bouton Envoyer vers base
    _AjouterBouton ws, "btn_Envoyer", "Envoyer vers base", _
        leftC, top19, 145, 34, CLR_VERT, "Mod_Comptabilite.EnvoyerVersBase"

    ' Bouton Mode d'emploi
    Dim leftD19 As Double : leftD19 = ws.Range("D19").Left
    _AjouterBouton ws, "btn_Manuel", "Mode d'emploi", _
        leftD19, top19, 130, 34, CLR_VIOLET, "Mod_Comptabilite.Ouvrir_Manuel"
End Sub

Private Sub _Creer_Bouton_DB(nom_feuille As String, action As String, label As String, _
                               cellule As String, couleur As Long, larg As Double, haut As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nom_feuille)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btn_" & nom_feuille Then shp.Delete
    Next shp
    _AjouterBouton ws, "btn_" & nom_feuille, label, _
        ws.Range(cellule).Left, ws.Range(cellule).Top, larg, haut, couleur, "Mod_Comptabilite." & action
End Sub

Private Sub _Creer_Bouton_Pivot()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PIVOT)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "btn_" Then shp.Delete
    Next shp
    _AjouterBouton ws, "btn_Rafraichir", "Rafraichir tableaux", _
        ws.Range("A20").Left, ws.Range("A20").Top, 180, 30, CLR_BLEU_FONCE, "Mod_Comptabilite.RafraichirTableaux"
End Sub

Private Sub _Creer_Bouton_TVA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TVA)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "btn_" Then shp.Delete
    Next shp
    _AjouterBouton ws, "btn_CalcTVA", "Recalculer TVA", _
        ws.Range("E3").Left, ws.Range("E3").Top, 160, 30, CLR_ROUGE, "Mod_Comptabilite.RecalculerTVA"
End Sub

' -- Fonction générique de création de bouton ----------------------------------

Private Sub _AjouterBouton(ws As Worksheet, nom As String, label As String, _
    Left As Double, Top As Double, larg As Double, haut As Double, _
    couleur As Long, action As String)

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, Left, Top, larg, haut)
    With shp
        .Name = nom
        .OnAction = action
        With .TextFrame
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment   = xlVAlignCenter
            With .Characters
                .Text = label
                .Font.Name  = "Calibri"
                .Font.Bold  = True
                .Font.Size  = 10
                .Font.Color = CLR_BLANC
            End With
        End With
        With .Fill
            .Visible = msoTrue
            .ForeColor.RGB = couleur
            .Solid
        End With
        .Line.Visible = msoFalse
        .Shadow.Visible = msoTrue
        .Shadow.Style  = msoShadowStyleInnerShadow
    End With
End Sub


' ============================================================
'  SECTION 2 — IMPORTATION PDF
' ============================================================

Public Sub ImporterPDF()
    Dim chemin As Variant
    chemin = Application.GetOpenFilename( _
        "Fichiers PDF (*.pdf), *.pdf", , _
        "Sélectionner une facture PDF")

    If chemin = False Then Exit Sub   ' annulation utilisateur

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PDF)

    ' Stocker chemin et réinitialiser les champs
    ws.Range("B5") = CStr(chemin)
    _Vider_Champs_Import ws
    _Statut ws, "Fichier sélectionné. Cliquez sur « Extraire données ».", "bleu"
End Sub

Private Sub _Vider_Champs_Import(ws As Worksheet)
    Dim i As Integer
    For i = 9 To 15
        ws.Cells(i, 2).Value = ""
    Next i
End Sub

Public Sub ExtraireViaPython()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PDF)

    Dim cheminPDF As String
    cheminPDF = Trim(ws.Range("B5").Value)

    If cheminPDF = "" Then
        MsgBox "Aucun fichier PDF sélectionné." & vbCrLf & _
               "Cliquez d'abord sur « Importer PDF ».", vbExclamation, "Fichier manquant"
        Exit Sub
    End If

    If Not _FichierExiste(cheminPDF) Then
        MsgBox "Le fichier n'existe plus :" & vbCrLf & cheminPDF, vbCritical, "Fichier introuvable"
        Exit Sub
    End If

    _Statut ws, "Extraction en cours… (patienter)", "orange"
    Application.StatusBar = "Extraction PDF en cours…"
    DoEvents

    ' Chemin JSON de sortie (dans %TEMP%)
    Dim jsonPath As String
    jsonPath = Environ("TEMP") & "\" & JSON_TEMP

    ' Exécuter Python
    Dim retour As Integer
    retour = _LancerPython(cheminPDF, jsonPath)

    If retour <> 0 Or Not _FichierExiste(jsonPath) Then
        _Statut ws, "Échec de l'extraction. Vérifiez que Python est installé (voir Guide).", "rouge"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Lire résultats JSON
    Dim json As String
    json = _LireTexte(jsonPath)

    Dim succes As Boolean
    succes = (_JSONVal(json, "succes") = "True" Or _JSONVal(json, "succes") = "true")

    If Not succes Then
        Dim msg As String
        msg = _JSONVal(json, "erreur")
        _Statut ws, "Erreur : " & msg, "rouge"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Remplir les champs de sortie
    ws.Range("B9").Value  = _JSONVal(json, "numero_facture")
    ws.Range("B10").Value = _JSONVal(json, "date_emission")
    ws.Range("B11").Value = CDbl_FR(_JSONVal(json, "montant_ht"))
    ws.Range("B12").Value = CDbl_FR(_JSONVal(json, "montant_tva"))
    ws.Range("B13").Value = CDbl_FR(_JSONVal(json, "montant_ttc"))
    ws.Range("B14").Value = CDbl_FR(_JSONVal(json, "taux_tva"))
    ws.Range("B15").Value = _JSONVal(json, "type_facture")

    ' Afficher avertissements éventuels
    Dim averts As String
    averts = _JSONVal(json, "avertissements")

    Dim msg2 As String
    msg2 = "Extraction réussie (" & _JSONVal(json, "moteur_pdf") & ")."
    If Len(averts) > 5 Then msg2 = msg2 & vbCrLf & "[!] " & averts
    _Statut ws, msg2, "vert"
    Application.StatusBar = False
End Sub

Public Sub EnvoyerVersBase()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PDF)

    ' Validation champs obligatoires
    If Trim(ws.Range("B9").Value) = "" Then
        MsgBox "Le numéro de facture est vide." & vbCrLf & _
               "Saisissez-le manuellement ou relancez l'extraction.", _
               vbExclamation, "Champ manquant"
        Exit Sub
    End If
    If CDbl_FR(ws.Range("B13").Value) = 0 Then
        If MsgBox("Le montant TTC est 0. Continuer quand même ?", _
                  vbQuestion + vbYesNo, "Montant nul") = vbNo Then Exit Sub
    End If

    Dim typeFacture As String
    typeFacture = Trim(ws.Range("B15").Value)
    If typeFacture = "" Then typeFacture = "Fournisseur"

    Dim cibleSheet As String
    If InStr(1, typeFacture, "Client", vbTextCompare) > 0 Then
        cibleSheet = SH_CLIENTS
    Else
        cibleSheet = SH_FOURN
    End If

    Dim wsDest As Worksheet
    Set wsDest = ThisWorkbook.Sheets(cibleSheet)

    ' Trouver la première ligne vide
    Dim derLigne As Long
    derLigne = _DerniereLigne(wsDest) + 1

    ' Générer ID facture
    Dim idFacture As String
    idFacture = _GenererID(wsDest, Left(cibleSheet, 1))  ' "F" ou "C"

    ' Écrire les données
    With wsDest
        .Cells(derLigne, 1).Value  = idFacture
        .Cells(derLigne, 2).Value  = ""                                   ' Nom fourn/client (à saisir)
        .Cells(derLigne, 3).Value  = ws.Range("B9").Value                 ' Numéro facture
        .Cells(derLigne, 4).Value  = _ConvertirDate(ws.Range("B10").Value)' Date émission
        .Cells(derLigne, 5).Value  = CDbl_FR(ws.Range("B11").Value)       ' HT
        .Cells(derLigne, 6).Value  = CDbl_FR(ws.Range("B12").Value)       ' TVA
        .Cells(derLigne, 7).Value  = CDbl_FR(ws.Range("B13").Value)       ' TTC
        .Cells(derLigne, 8).Value  = CDbl_FR(ws.Range("B14").Value)       ' Taux TVA
        .Cells(derLigne, 9).Value  = ""                                   ' Catégorie
        .Cells(derLigne, 10).Value = ws.Range("B5").Value                 ' Lien PDF
        .Cells(derLigne, 11).Value = "En attente"                         ' Statut paiement
        .Cells(derLigne, 12).Value = Now                                  ' Date import
        .Cells(derLigne, 13).Value = "Automatique"                        ' Source
        .Cells(derLigne, 4).NumberFormat  = "DD/MM/YYYY"
        .Cells(derLigne, 12).NumberFormat = "DD/MM/YYYY"
        Dim col As Variant
        For Each col In Array(5, 6, 7)
            .Cells(derLigne, col).NumberFormat = "#,##0.00 €"
        Next col
    End With

    _Statut ws, "[OK] Enregistre dans " & cibleSheet & " (ligne " & derLigne & ", ID : " & idFacture & ").", "vert"

    ' Proposer rafraîchissement
    If MsgBox("Données enregistrées." & vbCrLf & _
              "Rafraîchir les tableaux de bord maintenant ?", _
              vbQuestion + vbYesNo, "Mise à jour") = vbYes Then
        RafraichirTableaux
    End If
End Sub


' ============================================================
'  SECTION 3 — RAFRAÎCHISSEMENT ET TVA
' ============================================================

Public Sub RafraichirTableaux()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Forcer recalcul sur toutes les feuilles
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Calculate
    Next ws

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Tableaux de bord mis à jour.", vbInformation, "Rafraîchi"
End Sub

Public Sub RecalculerTVA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TVA)
    ws.Calculate
    MsgBox "Calcul TVA mis à jour pour l'année " & ws.Range("B3").Value & ".", _
           vbInformation, "TVA recalculée"
End Sub


' ============================================================
'  SECTION 4 — OUVERTURE PDF DEPUIS LES BASES
' ============================================================

Public Sub Ouvrir_PDF_Ligne()
    Dim ws As Worksheet
    Dim cell As Range

    ' Identifier la feuille active et la sélection
    Set ws = ActiveSheet
    If ws.Name <> SH_FOURN And ws.Name <> SH_CLIENTS Then
        MsgBox "Cette action fonctionne uniquement sur Fournisseurs_DB ou Clients_DB.", _
               vbExclamation, "Feuille incorrecte"
        Exit Sub
    End If

    ' Colonne J = 10 = Lien_PDF
    Dim ligneActive As Long
    ligneActive = ActiveCell.Row
    If ligneActive <= 1 Then
        MsgBox "Cliquez d'abord sur une ligne de données.", vbInformation, "Sélection"
        Exit Sub
    End If

    Dim cheminPDF As String
    cheminPDF = Trim(ws.Cells(ligneActive, 10).Value)

    If cheminPDF = "" Then
        MsgBox "Aucun lien PDF sur cette ligne.", vbExclamation, "Pas de PDF"
        Exit Sub
    End If

    If Not _FichierExiste(cheminPDF) Then
        MsgBox "Fichier introuvable :" & vbCrLf & cheminPDF, vbCritical, "Fichier manquant"
        Exit Sub
    End If

    Shell "cmd /c start """" """ & cheminPDF & """", vbHide
End Sub


Public Sub Ouvrir_Manuel()
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Manuel_EasyCompta.pdf"

    If Not _FichierExiste(pdfPath) Then
        MsgBox "Manuel introuvable :" & vbCrLf & pdfPath, vbExclamation, "Fichier manquant"
        Exit Sub
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c start """" """ & pdfPath & """", 0, False
End Sub


' ============================================================
'  SECTION 6 — UTILITAIRES INTERNES
' ============================================================

' Lance l'extracteur PDF et attend la fin.
' Priorité 1 : pdf_extractor.exe (version installée via MSI, sans Python requis)
' Priorité 2 : pdf_extractor.py via Python (mode développement)
Private Function _LancerPython(cheminPDF As String, jsonSortie As String) As Integer
    Dim scriptDir As String
    scriptDir = ThisWorkbook.Path

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim cmd As String

    ' -- Priorité 1 : exe autonome --
    Dim exePath As String
    exePath = scriptDir & "\pdf_extractor.exe"
    If _FichierExiste(exePath) Then
        cmd = """" & exePath & """ """ & cheminPDF & """ """ & jsonSortie & """"
        _LancerPython = wsh.Run("cmd /c " & cmd, 0, True)
        Exit Function
    End If

    ' -- Priorité 2 : script Python (mode développement) --
    Dim python As String
    python = _TrouverPython()
    If python = "" Then
        MsgBox "Extracteur PDF introuvable." & vbCrLf & vbCrLf & _
               "Solution 1 : Installez EasyCompta via l'installateur .msi" & vbCrLf & _
               "Solution 2 : Installez Python 3.10+ et cochez « Add Python to PATH ».", _
               vbCritical, "Extracteur manquant"
        _LancerPython = 1
        Exit Function
    End If

    cmd = """" & python & """ """ & scriptDir & "\" & SCRIPT_PDF & """ " & _
          """" & cheminPDF & """ """ & jsonSortie & """"
    _LancerPython = wsh.Run("cmd /c " & cmd, 0, True)
End Function

Private Function _TrouverPython() As String
    Dim candidats As Variant
    candidats = Array("python", "python3", _
        "C:\Python313\python.exe", "C:\Python312\python.exe", _
        "C:\Python311\python.exe", "C:\Python310\python.exe", _
        Environ("LOCALAPPDATA") & "\Programs\Python\Python313\python.exe", _
        Environ("LOCALAPPDATA") & "\Programs\Python\Python312\python.exe", _
        Environ("LOCALAPPDATA") & "\Programs\Python\Python311\python.exe", _
        Environ("LOCALAPPDATA") & "\Programs\Python\Python310\python.exe")

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim c As Variant
    For Each c In candidats
        If _FichierExiste(CStr(c)) Or _TestCommande(CStr(c)) Then
            _TrouverPython = CStr(c)
            Exit Function
        End If
    Next c
    _TrouverPython = ""
End Function

Private Function _TestCommande(cmd As String) As Boolean
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    _TestCommande = (wsh.Run("cmd /c " & cmd & " --version >nul 2>&1", 0, True) = 0)
    On Error GoTo 0
End Function

Private Function _FichierExiste(chemin As String) As Boolean
    _FichierExiste = (Len(Dir(chemin)) > 0)
End Function

' Lecture d'un fichier texte complet
Private Function _LireTexte(chemin As String) As String
    Dim num As Integer
    num = FreeFile
    Open chemin For Input As #num
    _LireTexte = Input(LOF(num), num)
    Close #num
End Function

' Extrait la valeur d'une clé dans un JSON simple (pas de tableaux imbriqués)
Private Function _JSONVal(json As String, cle As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Cherche "cle": "valeur"
    re.Pattern = """" & cle & """\s*:\s*""([^""]*)"""
    re.IgnoreCase = True
    If re.Test(json) Then
        _JSONVal = re.Execute(json)(0).SubMatches(0)
        Exit Function
    End If

    ' Cherche "cle": valeur_numerique_ou_booleen
    re.Pattern = """" & cle & """\s*:\s*([^,\}\[\n]+)"
    If re.Test(json) Then
        _JSONVal = Trim(re.Execute(json)(0).SubMatches(0))
        Exit Function
    End If

    _JSONVal = ""
End Function

' Convertit un string en Double (supporte , et . comme séparateurs)
Private Function CDbl_FR(valeur As Variant) As Double
    On Error Resume Next
    Dim s As String
    s = Trim(CStr(valeur))
    s = Replace(s, " ", "")
    ' Si les deux séparateurs présents, format 1.234,56
    If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
        If InStr(s, ".") < InStr(s, ",") Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    ElseIf InStr(s, ",") > 0 Then
        s = Replace(s, ",", ".")
    End If
    CDbl_FR = CDbl(s)
    If Err.Number <> 0 Then CDbl_FR = 0
    On Error GoTo 0
End Function

' Convertit une date texte JJ/MM/AAAA en Date Excel
Private Function _ConvertirDate(valeur As String) As Variant
    On Error Resume Next
    valeur = Trim(valeur)
    If valeur = "" Then _ConvertirDate = Empty : Exit Function
    valeur = Replace(valeur, "-", "/")
    valeur = Replace(valeur, ".", "/")
    Dim parties() As String
    parties = Split(valeur, "/")
    If UBound(parties) = 2 Then
        _ConvertirDate = DateSerial(CInt(parties(2)), CInt(parties(1)), CInt(parties(0)))
    Else
        _ConvertirDate = CDate(valeur)
    End If
    If Err.Number <> 0 Then _ConvertirDate = Empty
    On Error GoTo 0
End Function

' Génère un ID unique : F2025-0001 ou C2025-0001
Private Function _GenererID(ws As Worksheet, prefixe As String) As String
    Dim der As Long
    der = _DerniereLigne(ws)
    Dim annee As String
    annee = CStr(Year(Now))
    If der <= 1 Then
        _GenererID = prefixe & annee & "-0001"
    Else
        _GenererID = prefixe & annee & "-" & Format(der, "0000")
    End If
End Function

' Retourne le numéro de la dernière ligne avec données en colonne A
Private Function _DerniereLigne(ws As Worksheet) As Long
    _DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

' Affiche un message de statut coloré dans la zone statut de PDF_Import
Private Sub _Statut(ws As Worksheet, message As String, niveau As String)
    Dim c As Range
    Set c = ws.Range("A23")
    c.Value = message
    c.WrapText = True
    Select Case LCase(niveau)
        Case "vert"   : c.Font.Color = RGB(56, 87, 35)   : c.Interior.Color = RGB(198, 224, 180)
        Case "orange" : c.Font.Color = RGB(156, 87, 0)   : c.Interior.Color = RGB(255, 230, 153)
        Case "rouge"  : c.Font.Color = RGB(192, 0, 0)    : c.Interior.Color = RGB(255, 199, 206)
        Case Else     : c.Font.Color = RGB(31, 78, 121)  : c.Interior.Color = RGB(189, 215, 238)
    End Select
    DoEvents
End Sub


' ============================================================
'  SECTION 7 - LICENCE
' ============================================================

Private Const SECRET_LICENCE As String = "SEB$C0mpt4bl3#Pr0j3t9!"
Private Const REG_PATH        As String = "HKCU\Software\EasyCompta\"

Public Sub VerifierOuDemanderLicence()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim email As String, cle As String
    On Error Resume Next
    email = wsh.RegRead(REG_PATH & "Email")
    cle   = wsh.RegRead(REG_PATH & "Cle")
    On Error GoTo 0

    If email <> "" And cle <> "" And _VerifierLicence(cle, email) Then
        Exit Sub
    End If

    Dim rep As Integer
    rep = MsgBox("Merci d'utiliser EasyCompta !" & vbCrLf & vbCrLf & _
                 "Ce logiciel est libre et gratuit." & vbCrLf & _
                 "Si vous souhaitez soutenir le projet, un don PayPal" & vbCrLf & _
                 "vous permet d'obtenir une cle d'activation." & vbCrLf & vbCrLf & _
                 "Voulez-vous entrer votre cle de licence maintenant ?", _
                 vbQuestion + vbYesNo, "Activation EasyCompta")

    If rep = vbNo Then Exit Sub

    Do
        email = InputBox("Entrez votre adresse email :", "Activation", email)
        If email = "" Then Exit Sub

        cle = InputBox("Entrez votre cle de licence :" & vbCrLf & _
                       "(format : SEB-AAAA-XXXX-XXXX)", "Activation", cle)
        If cle = "" Then Exit Sub

        If _VerifierLicence(cle, email) Then
            wsh.RegWrite REG_PATH & "Email", LCase(Trim(email)), "REG_SZ"
            wsh.RegWrite REG_PATH & "Cle",   UCase(Trim(cle)),   "REG_SZ"
            MsgBox "Licence activee avec succes !" & vbCrLf & "Merci pour votre soutien.", _
                   vbInformation, "Activation reussie"
            Exit Sub
        Else
            MsgBox "Cle invalide pour cet email." & vbCrLf & _
                   "Verifiez l'email et la cle, puis reessayez.", _
                   vbCritical, "Cle incorrecte"
        End If
    Loop
End Sub

Private Function _VerifierLicence(cle As String, email As String) As Boolean
    cle = UCase(Trim(cle))
    If Len(cle) <> 18 Then Exit Function
    If Left$(cle, 4) <> "SEB-" Then Exit Function

    Dim hashCle As String
    hashCle = Mid$(cle, 10, 4) & Mid$(cle, 15, 4)

    Dim expected As String
    expected = _HashCle(LCase(Trim(email)) & SECRET_LICENCE)

    _VerifierLicence = (hashCle = Left$(expected, 8))
End Function

Private Function _HashCle(texte As String) As String
    Dim h1 As Double, h2 As Double
    Dim m1 As Double, m2 As Double
    Dim i As Long, v As Long
    h1 = 1505843#
    h2 = 1234567891#
    For i = 1 To Len(texte)
        v = Asc(Mid$(texte, i, 1))
        m1 = h1 - Int(h1 / 67108864#) * 67108864#
        m2 = h2 - Int(h2 / 67108864#) * 67108864#
        h1 = m1 * 31# + v
        h2 = m2 * 37# + v
        h1 = h1 - Int(h1 / 2147483648#) * 2147483648#
        h2 = h2 - Int(h2 / 2147483648#) * 2147483648#
    Next i
    _HashCle = Right$("00000000" & Hex$(CLng(h1) Xor CLng(h2)), 8)
End Function
