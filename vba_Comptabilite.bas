Attribute VB_Name = "Mod_Comptabilite"
' ============================================================
'  MODULE : Mod_Comptabilite
'  Projet : Système de gestion comptable ECPT
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

' Licence
Private Const REG_PATH        As String = "HKCU\Software\EasyCompta"
Private Const APP_VERSION     As String = "2.0.0"


' ============================================================
'  SECTION 1 - SETUP INITIAL
' ============================================================

Public Sub Setup_Boutons()
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    Call Creer_Boutons_PDF_Import
    Call Creer_Bouton_DB(SH_FOURN,   "Ouvrir_PDF_Ligne",  "Ouvrir PDF",       "D1", CLR_BLEU_FONCE, 120, 26)
    Call Creer_Bouton_DB(SH_CLIENTS, "Ouvrir_PDF_Ligne",  "Ouvrir PDF",       "D1", CLR_VERT,        120, 26)
    Call Creer_Bouton_DB2(SH_FOURN,   "MettreAJourStatut", "Statut paiement", "F1", CLR_ORANGE, 140, 26)
    Call Creer_Bouton_DB2(SH_CLIENTS, "MettreAJourStatut", "Statut paiement", "F1", CLR_ORANGE, 140, 26)
    Call Creer_Bouton_DB3(SH_CLIENTS, "GenererRelance",    "Relance",         "H1", CLR_ROUGE,  120, 26)
    Call Creer_Bouton_Pivot
    Call Creer_Bouton_TVA
    ColoriserLignesDB ThisWorkbook.Sheets(SH_FOURN)
    ColoriserLignesDB ThisWorkbook.Sheets(SH_CLIENTS)

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

Private Sub Creer_Boutons_PDF_Import()
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
    AjouterBouton ws, "btn_ImporterPDF", "Importer PDF", _
        leftD5, topD5, 140, 26, CLR_BLEU_FONCE, "Mod_Comptabilite.ImporterPDF"

    ' Bouton Extraire données
    AjouterBouton ws, "btn_Extraire", "Extraire donnees", _
        leftB, top19, 145, 34, CLR_BLEU_MOY, "Mod_Comptabilite.ExtraireViaPython"

    ' Bouton Envoyer vers base
    AjouterBouton ws, "btn_Envoyer", "Envoyer vers base", _
        leftC, top19, 145, 34, CLR_VERT, "Mod_Comptabilite.EnvoyerVersBase"

    ' Bouton Mode d'emploi
    Dim leftD19 As Double : leftD19 = ws.Range("D19").Left
    AjouterBouton ws, "btn_Manuel", "Mode d'emploi", _
        leftD19, top19, 130, 34, CLR_VIOLET, "Mod_Comptabilite.Ouvrir_Manuel"
End Sub

Private Sub Creer_Bouton_DB(nom_feuille As String, action As String, label As String, _
                               cellule As String, couleur As Long, larg As Double, haut As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nom_feuille)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btn_" & nom_feuille Then shp.Delete
    Next shp
    AjouterBouton ws, "btn_" & nom_feuille, label, _
        ws.Range(cellule).Left, ws.Range(cellule).Top, larg, haut, couleur, "Mod_Comptabilite." & action
End Sub

Private Sub Creer_Bouton_DB2(nom_feuille As String, action As String, label As String, _
                                cellule As String, couleur As Long, larg As Double, haut As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nom_feuille)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btn2_" & nom_feuille Then shp.Delete
    Next shp
    AjouterBouton ws, "btn2_" & nom_feuille, label, _
        ws.Range(cellule).Left, ws.Range(cellule).Top, larg, haut, couleur, "Mod_Comptabilite." & action
End Sub

Private Sub Creer_Bouton_DB3(nom_feuille As String, action As String, label As String, _
                                cellule As String, couleur As Long, larg As Double, haut As Double)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nom_feuille)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btn3_" & nom_feuille Then shp.Delete
    Next shp
    AjouterBouton ws, "btn3_" & nom_feuille, label, _
        ws.Range(cellule).Left, ws.Range(cellule).Top, larg, haut, couleur, "Mod_Comptabilite." & action
End Sub

Private Sub Creer_Bouton_Pivot()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_PIVOT)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "btn_" Then shp.Delete
    Next shp
    AjouterBouton ws, "btn_Rafraichir", "Rafraichir tableaux", _
        ws.Range("A20").Left, ws.Range("A20").Top, 180, 30, CLR_BLEU_FONCE, "Mod_Comptabilite.RafraichirTableaux"
End Sub

Private Sub Creer_Bouton_TVA()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TVA)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "btn_" Then shp.Delete
    Next shp
    AjouterBouton ws, "btn_CalcTVA", "Recalculer TVA", _
        ws.Range("E3").Left, ws.Range("E3").Top, 160, 30, CLR_ROUGE, "Mod_Comptabilite.RecalculerTVA"
    AjouterBouton ws, "btn_ExportCSV", "Exporter CSV", _
        ws.Range("H3").Left, ws.Range("H3").Top, 160, 30, CLR_VERT, "Mod_Comptabilite.ExporterCSV"
End Sub

' -- Fonction générique de création de bouton ----------------------------------

Private Sub AjouterBouton(ws As Worksheet, nom As String, label As String, _
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
'  SECTION 2 - IMPORTATION PDF
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
    Vider_Champs_Import ws
    Statut ws, "Fichier sélectionné. Cliquez sur « Extraire données ».", "bleu"
End Sub

Private Sub Vider_Champs_Import(ws As Worksheet)
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

    If Not FichierExiste(cheminPDF) Then
        MsgBox "Le fichier n'existe plus :" & vbCrLf & cheminPDF, vbCritical, "Fichier introuvable"
        Exit Sub
    End If

    Statut ws, "Extraction en cours… (patienter)", "orange"
    Application.StatusBar = "Extraction PDF en cours…"
    DoEvents

    ' Chemin JSON de sortie (dans %TEMP%)
    Dim jsonPath As String
    jsonPath = Environ("TEMP") & "\" & JSON_TEMP

    ' Exécuter Python
    Dim retour As Integer
    retour = LancerPython(cheminPDF, jsonPath)

    If retour <> 0 Or Not FichierExiste(jsonPath) Then
        Statut ws, "Échec de l'extraction. Vérifiez que Python est installé (voir Guide).", "rouge"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Lire résultats JSON
    Dim json As String
    json = LireTexte(jsonPath)

    Dim succes As Boolean
    succes = (JSONVal(json, "succes") = "True" Or JSONVal(json, "succes") = "true")

    If Not succes Then
        Dim msg As String
        msg = JSONVal(json, "erreur")
        Statut ws, "Erreur : " & msg, "rouge"
        Application.StatusBar = False
        Exit Sub
    End If

    ' Remplir les champs de sortie
    ws.Range("B9").Value  = JSONVal(json, "numero_facture")
    ws.Range("B10").Value = JSONVal(json, "date_emission")
    ws.Range("B11").Value = CDbl_FR(JSONVal(json, "montant_ht"))
    ws.Range("B12").Value = CDbl_FR(JSONVal(json, "montant_tva"))
    ws.Range("B13").Value = CDbl_FR(JSONVal(json, "montant_ttc"))
    ws.Range("B14").Value = CDbl_FR(JSONVal(json, "taux_tva"))
    ws.Range("B15").Value = JSONVal(json, "type_facture")

    ' Afficher avertissements éventuels
    Dim averts As String
    averts = JSONVal(json, "avertissements")

    Dim msg2 As String
    msg2 = "Extraction réussie (" & JSONVal(json, "moteur_pdf") & ")."
    If Len(averts) > 5 Then msg2 = msg2 & vbCrLf & "[!] " & averts
    Statut ws, msg2, "vert"
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
    derLigne = DerniereLigne(wsDest) + 1

    ' Générer ID facture
    Dim idFacture As String
    idFacture = GenererID(wsDest, Left(cibleSheet, 1))  ' "F" ou "C"

    ' Écrire les données
    With wsDest
        .Cells(derLigne, 1).Value  = idFacture
        .Cells(derLigne, 2).Value  = ""                                   ' Nom fourn/client (à saisir)
        .Cells(derLigne, 3).Value  = ws.Range("B9").Value                 ' Numéro facture
        .Cells(derLigne, 4).Value  = ConvertirDate(ws.Range("B10").Value)' Date émission
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

    Statut ws, "[OK] Enregistre dans " & cibleSheet & " (ligne " & derLigne & ", ID : " & idFacture & ").", "vert"
    ColoriserLigne wsDest, derLigne

    ' Proposer rafraîchissement
    If MsgBox("Données enregistrées." & vbCrLf & _
              "Rafraîchir les tableaux de bord maintenant ?", _
              vbQuestion + vbYesNo, "Mise à jour") = vbYes Then
        RafraichirTableaux
    End If
End Sub


' ============================================================
'  SECTION 3 - RAFRAICHISSEMENT ET TVA
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
'  SECTION 4 - OUVERTURE PDF DEPUIS LES BASES
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
    If ligneActive <= 2 Then
        MsgBox "Cliquez d'abord sur une ligne de données.", vbInformation, "Sélection"
        Exit Sub
    End If

    Dim cheminPDF As String
    cheminPDF = Trim(ws.Cells(ligneActive, 10).Value)

    If cheminPDF = "" Then
        MsgBox "Aucun lien PDF sur cette ligne.", vbExclamation, "Pas de PDF"
        Exit Sub
    End If

    If Not FichierExiste(cheminPDF) Then
        MsgBox "Fichier introuvable :" & vbCrLf & cheminPDF, vbCritical, "Fichier manquant"
        Exit Sub
    End If

    Shell "cmd /c start """" """ & cheminPDF & """", vbHide
End Sub


Public Sub Ouvrir_Manuel()
    Dim pdfPath As String
    pdfPath = ThisWorkbook.Path & "\Manuel_EasyCompta.pdf"

    If Not FichierExiste(pdfPath) Then
        MsgBox "Manuel introuvable :" & vbCrLf & pdfPath, vbExclamation, "Fichier manquant"
        Exit Sub
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c start """" """ & pdfPath & """", 0, False
End Sub


' ============================================================
'  SECTION 4b - GESTION STATUT PAIEMENT
' ============================================================

Public Sub MettreAJourStatut()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.Name <> SH_FOURN And ws.Name <> SH_CLIENTS Then
        MsgBox "Cette action fonctionne uniquement sur Fournisseurs_DB ou Clients_DB.", _
               vbExclamation, "Feuille incorrecte"
        Exit Sub
    End If

    Dim ligne As Long
    ligne = ActiveCell.Row
    If ligne <= 2 Or Trim(CStr(ws.Cells(ligne, 1).Value)) = "" Then
        MsgBox "S" & Chr(233) & "lectionnez d'abord une ligne de donn" & Chr(233) & "es.", _
               vbInformation, "S" & Chr(233) & "lection"
        Exit Sub
    End If

    Dim numFacture As String : numFacture  = CStr(ws.Cells(ligne, 3).Value)
    Dim statutActuel As String : statutActuel = CStr(ws.Cells(ligne, 11).Value)

    Dim rep As Integer
    rep = MsgBox("Facture : " & numFacture & vbCrLf & _
                 "Statut actuel : " & statutActuel & vbCrLf & vbCrLf & _
                 "[Oui]     " & ChrW(8594) & " Marquer comme Pay" & Chr(233) & vbCrLf & _
                 "[Non]     " & ChrW(8594) & " Autre statut" & vbCrLf & _
                 "[Annuler] " & ChrW(8594) & " Ne rien changer", _
                 vbQuestion + vbYesNoCancel, "Statut paiement")

    If rep = vbCancel Then Exit Sub

    Dim nouveauStatut As String
    If rep = vbYes Then
        nouveauStatut = "Pay" & Chr(233)
    Else
        Dim choix As String
        choix = InputBox("Choisissez le nouveau statut :" & vbCrLf & vbCrLf & _
                         "1 - En attente" & vbCrLf & _
                         "2 - En retard" & vbCrLf & _
                         "3 - Annul" & Chr(233), _
                         "Statut paiement", "1")
        If choix = "" Then Exit Sub
        Select Case Trim(choix)
            Case "1" : nouveauStatut = "En attente"
            Case "2" : nouveauStatut = "En retard"
            Case "3" : nouveauStatut = "Annul" & Chr(233)
            Case Else
                MsgBox "Choix invalide (1, 2 ou 3).", vbExclamation
                Exit Sub
        End Select
    End If

    ws.Cells(ligne, 11).Value = nouveauStatut

    If nouveauStatut = "Pay" & Chr(233) Then
        Dim dateStr As String
        dateStr = InputBox("Date de r" & Chr(232) & "glement :", _
                           "R" & Chr(232) & "glement", Format(Date, "DD/MM/YYYY"))
        If dateStr <> "" Then
            Dim dt As Variant
            dt = ConvertirDate(dateStr)
            If Not IsEmpty(dt) Then
                ws.Cells(ligne, 14).Value = dt
                ws.Cells(ligne, 14).NumberFormat = "DD/MM/YYYY"
            End If
        End If
    Else
        ws.Cells(ligne, 14).Value = ""
    End If

    ColoriserLigne ws, ligne
End Sub

Private Sub ColoriserLigne(ws As Worksheet, ligne As Long)
    Dim statut As String
    statut = Trim(CStr(ws.Cells(ligne, 11).Value))
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(ligne, 1), ws.Cells(ligne, 15))
    Select Case statut
        Case "Pay" & Chr(233)
            rng.Interior.Color = RGB(198, 239, 206)
            rng.Font.Color     = RGB(0, 97, 0)
        Case "En attente"
            rng.Interior.Color = RGB(255, 235, 156)
            rng.Font.Color     = RGB(156, 101, 0)
        Case "En retard"
            rng.Interior.Color = RGB(255, 199, 206)
            rng.Font.Color     = RGB(156, 0, 6)
        Case "Annul" & Chr(233)
            rng.Interior.Color = RGB(217, 217, 217)
            rng.Font.Color     = RGB(89, 89, 89)
        Case Else
            rng.Interior.Pattern = xlNone
            rng.Font.Color       = RGB(0, 0, 0)
    End Select
End Sub

Private Sub ColoriserLignesDB(ws As Worksheet)
    Dim der As Long : der = DerniereLigne(ws)
    If der <= 2 Then Exit Sub
    Dim i As Long
    For i = 3 To der
        If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
            ColoriserLigne ws, i
        End If
    Next i
End Sub


' ============================================================
'  SECTION 4c - RELANCES CLIENTS
' ============================================================

Public Sub GenererRelance()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    If ws.Name <> SH_CLIENTS Then
        MsgBox "Les relances ne s'appliquent qu'aux Clients_DB.", _
               vbExclamation, "Feuille incorrecte"
        Exit Sub
    End If

    Dim ligne As Long
    ligne = ActiveCell.Row
    If ligne <= 2 Or Trim(CStr(ws.Cells(ligne, 1).Value)) = "" Then
        MsgBox "S" & Chr(233) & "lectionnez d'abord une ligne de donn" & Chr(233) & "es.", _
               vbInformation, "S" & Chr(233) & "lection"
        Exit Sub
    End If

    Dim statut As String : statut = Trim(CStr(ws.Cells(ligne, 11).Value))
    If statut <> "En retard" Then
        MsgBox "La relance ne s'applique qu'aux factures ""En retard""." & vbCrLf & _
               "Statut actuel : " & statut, vbExclamation, "Statut incorrect"
        Exit Sub
    End If

    ' Incrémenter le compteur de relances (col 15)
    Dim nbRel As Long
    nbRel = 0
    On Error Resume Next
    nbRel = CLng(ws.Cells(ligne, 15).Value)
    On Error GoTo 0
    nbRel = nbRel + 1
    ws.Cells(ligne, 15).Value = nbRel

    ' Récupérer les données de la ligne
    Dim nomClient  As String : nomClient  = Trim(CStr(ws.Cells(ligne, 2).Value))
    Dim numFacture As String : numFacture = Trim(CStr(ws.Cells(ligne, 3).Value))
    Dim dateFacture As String : dateFacture = ""
    On Error Resume Next
    If Not IsEmpty(ws.Cells(ligne, 4).Value) And ws.Cells(ligne, 4).Value <> 0 Then
        dateFacture = Format(CDate(ws.Cells(ligne, 4).Value), "DD/MM/YYYY")
    End If
    On Error GoTo 0
    Dim montantTTC As Double : montantTTC = 0
    On Error Resume Next
    montantTTC = CDbl(ws.Cells(ligne, 7).Value)
    On Error GoTo 0

    ' Écrire le HTML dans %TEMP% et l'ouvrir
    Dim htmlPath As String
    htmlPath = Environ("TEMP") & "\relance_" & Format(Now, "YYYYMMDD_HHMMSS") & ".html"
    Dim html As String
    html = HtmlRelance(nomClient, numFacture, dateFacture, montantTTC, nbRel)

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "utf-8"
    ado.Open
    ado.WriteText html
    ado.SaveToFile htmlPath, 2
    ado.Close
    Set ado = Nothing

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd /c start """" """ & htmlPath & """", 0, False

    MsgBox OrdinalRelance(nbRel) & " relance g" & Chr(233) & "n" & Chr(233) & "r" & Chr(233) & _
           "e pour " & nomClient & "." & vbCrLf & _
           "Imprimez ou enregistrez en PDF depuis votre navigateur.", _
           vbInformation, "Relance client"
End Sub

Private Function OrdinalRelance(n As Long) As String
    If n = 1 Then
        OrdinalRelance = "1" & Chr(232) & "re"
    Else
        OrdinalRelance = CStr(n) & Chr(232) & "me"
    End If
End Function

Private Function HtmlRelance(nomClient As String, numFacture As String, _
                              dateFacture As String, montantTTC As Double, _
                              nbRel As Long) As String
    Dim montantStr As String
    montantStr = Replace(Format(montantTTC, "0.00"), ".", ",") & " " & ChrW(8364)
    Dim dateAujourd As String : dateAujourd = Format(Date, "DD/MM/YYYY")

    Dim objet As String, intro As String, corps As String, conclusion As String
    Select Case nbRel
        Case 1
            objet = "Rappel de facture impay" & Chr(233) & "e n" & Chr(176) & " " & numFacture
            intro = "Nous vous contactons afin de vous rappeler que la facture mentionn" & Chr(233) & _
                    "e ci-dessous reste " & Chr(224) & " ce jour sans r" & Chr(232) & "glement."
            corps = "Nous vous remercions de bien vouloir proc" & Chr(233) & _
                    "der au paiement dans les plus brefs d" & Chr(233) & "lais."
            conclusion = "Dans l'attente de votre r" & Chr(232) & "glement, nous restons " & _
                         Chr(224) & " votre disposition."
        Case 2
            objet = "2" & Chr(232) & "me relance — Facture impay" & Chr(233) & "e n" & Chr(176) & " " & numFacture
            intro = "Malgr" & Chr(233) & " notre premier rappel, la facture ci-dessous demeure impay" & Chr(233) & "e."
            corps = "Nous vous demandons instamment de r" & Chr(233) & "gler cette somme sous 8 jours ouvrables."
            conclusion = "Sans r" & Chr(233) & "glement dans ce d" & Chr(233) & "lai, nous nous verrons contraints " & _
                         "d'engager une proc" & Chr(233) & "dure de recouvrement."
        Case Else
            objet = "Mise en demeure — Facture n" & Chr(176) & " " & numFacture
            intro = "Suite " & Chr(224) & " vos pr" & Chr(233) & "c" & Chr(233) & "dentes relances rest" & Chr(233) & _
                    "es sans r" & Chr(233) & "ponse, nous constatons le non-paiement persistant de la facture ci-dessous."
            corps = "Nous vous mettons en demeure de r" & Chr(233) & "gler la somme due dans un d" & Chr(233) & _
                    "lai de 5 jours ouvrables."
            conclusion = Chr(192) & " d" & Chr(233) & "faut de paiement, nous proc" & Chr(233) & "derons " & _
                         Chr(224) & " un recouvrement judiciaire sans autre pr" & Chr(233) & "avis."
    End Select

    Dim css As String
    css = "body{font-family:Arial,sans-serif;font-size:13px;color:#222;margin:40px;max-width:780px}"
    css = css & "h1{font-size:18px;color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:6px}"
    css = css & ".header{display:flex;justify-content:space-between;margin-bottom:30px}"
    css = css & ".company{font-weight:bold;font-size:15px;color:#1F4E79}"
    css = css & ".date{color:#555;font-size:12px}"
    css = css & "table{width:100%;border-collapse:collapse;margin:20px 0}"
    css = css & "th{background:#1F4E79;color:#fff;padding:8px;text-align:left}"
    css = css & "td{padding:8px;border-bottom:1px solid #ddd}"
    css = css & ".montant{font-weight:bold;color:#C00000;font-size:16px}"
    css = css & ".footer{margin-top:40px;font-size:11px;color:#888;border-top:1px solid #ccc;padding-top:10px}"
    css = css & "@media print{body{margin:15px}}"

    Dim thRow As String
    thRow = "<tr><th>N" & Chr(176) & " Facture</th>"
    thRow = thRow & "<th>Date d'" & Chr(233) & "mission</th>"
    thRow = thRow & "<th>Montant TTC</th></tr>"

    Dim tdRow As String
    tdRow = "<tr><td>" & numFacture & "</td><td>" & dateFacture & "</td>"
    tdRow = tdRow & "<td class=""montant"">" & montantStr & "</td></tr>"

    Dim footerTxt As String
    footerTxt = "G" & Chr(233) & "n" & Chr(233) & "r" & Chr(233)
    footerTxt = footerTxt & " par EasyCompta v" & APP_VERSION & " le " & dateAujourd & "."

    Dim h As String
    h = "<!DOCTYPE html><html lang=""fr""><head><meta charset=""UTF-8"">"
    h = h & "<style>" & css & "</style></head><body>"
    h = h & "<div class=""header"">"
    h = h & "<div class=""company"">EasyCompta</div>"
    h = h & "<div class=""date"">Le " & dateAujourd & "</div></div>"
    h = h & "<h1>Lettre de relance — " & OrdinalRelance(nbRel) & " relance</h1>"
    h = h & "<p><strong>Destinataire :</strong> " & nomClient & "</p>"
    h = h & "<p><strong>Objet :</strong> " & objet & "</p>"
    h = h & "<p>" & intro & "</p>"
    h = h & "<table>" & thRow & tdRow & "</table>"
    h = h & "<p>" & corps & "</p>"
    h = h & "<p>" & conclusion & "</p>"
    h = h & "<p style=""margin-top:40px"">Cordialement,<br><strong>EasyCompta</strong></p>"
    h = h & "<div class=""footer"">" & footerTxt & "</div>"
    h = h & "</body></html>"
    HtmlRelance = h
End Function


' ============================================================
'  SECTION 5 - EXPORT CSV COMPTABLE
' ============================================================

Public Sub ExporterCSV()
    Dim annee  As String : annee  = CStr(Year(Now))
    Dim mois   As String : mois   = Format(Month(Now), "00")
    Dim nom    As String : nom    = "EasyCompta_export_" & annee & "-" & mois & ".csv"
    Dim chemin As String : chemin = ThisWorkbook.Path & "\" & nom

    If MsgBox("Exporter toutes les factures en CSV ?" & vbCrLf & vbCrLf & _
              "Fichier : " & nom & vbCrLf & _
              "Dossier : " & ThisWorkbook.Path, _
              vbQuestion + vbYesNo, "Export CSV") = vbNo Then Exit Sub

    Dim SEP As String : SEP = ";"
    Dim lignes As String
    lignes = "Type" & SEP & "N" & Chr(176) & " Facture" & SEP & "Date" & SEP & "Nom" & SEP & _
             "Montant HT" & SEP & "Montant TVA" & SEP & "Montant TTC" & SEP & _
             "Taux TVA %" & SEP & "Cat" & Chr(233) & "gorie" & SEP & "Statut" & SEP & _
             "R" & Chr(233) & "f" & Chr(233) & "rence PDF" & vbCrLf
    lignes = lignes & CsvLignesFeuille(SH_FOURN, "Fournisseur", SEP)
    lignes = lignes & CsvLignesFeuille(SH_CLIENTS, "Client", SEP)

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2
    ado.Charset = "utf-8"
    ado.Open
    ado.WriteText lignes
    ado.SaveToFile chemin, 2
    ado.Close
    Set ado = Nothing

    If MsgBox("Export termin" & Chr(233) & "." & vbCrLf & nom & " cr" & Chr(233) & Chr(233) & _
              " dans :" & vbCrLf & ThisWorkbook.Path & vbCrLf & vbCrLf & "Ouvrir le fichier ?", _
              vbQuestion + vbYesNo, "Export CSV") = vbYes Then
        Dim wsh As Object
        Set wsh = CreateObject("WScript.Shell")
        wsh.Run "cmd /c start """" """ & chemin & """", 0, False
    End If
End Sub

Private Function CsvLignesFeuille(nomFeuille As String, typeDoc As String, SEP As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(nomFeuille)
    Dim der As Long : der = DerniereLigne(ws)
    Dim result As String : result = ""
    Dim i As Long
    For i = 3 To der
        If Trim(CStr(ws.Cells(i, 1).Value)) <> "" Then
            result = result & _
                typeDoc & SEP & _
                CsvVal(CStr(ws.Cells(i, 3).Value)) & SEP & _
                CsvDate(ws.Cells(i, 4).Value) & SEP & _
                CsvVal(CStr(ws.Cells(i, 2).Value)) & SEP & _
                CsvMontant(ws.Cells(i, 5).Value) & SEP & _
                CsvMontant(ws.Cells(i, 6).Value) & SEP & _
                CsvMontant(ws.Cells(i, 7).Value) & SEP & _
                CsvMontant(ws.Cells(i, 8).Value) & SEP & _
                CsvVal(CStr(ws.Cells(i, 9).Value)) & SEP & _
                CsvVal(CStr(ws.Cells(i, 11).Value)) & SEP & _
                CsvVal(CStr(ws.Cells(i, 10).Value)) & vbCrLf
        End If
    Next i
    CsvLignesFeuille = result
End Function

Private Function CsvVal(v As String) As String
    v = Trim(v)
    If InStr(v, ";") > 0 Or InStr(v, """") > 0 Or InStr(v, Chr(10)) > 0 Then
        CsvVal = """" & Replace(v, """", """""") & """"
    Else
        CsvVal = v
    End If
End Function

Private Function CsvDate(v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Or CStr(v) = "" Or v = 0 Then
        CsvDate = ""
    Else
        CsvDate = Format(CDate(v), "DD/MM/YYYY")
    End If
    If Err.Number <> 0 Then CsvDate = CStr(v)
    On Error GoTo 0
End Function

Private Function CsvMontant(v As Variant) As String
    On Error Resume Next
    If IsEmpty(v) Or CStr(v) = "" Then
        CsvMontant = "0,00"
    Else
        CsvMontant = Replace(Format(CDbl_FR(v), "0.00"), ".", ",")
    End If
    If Err.Number <> 0 Then CsvMontant = "0,00"
    On Error GoTo 0
End Function


' ============================================================
'  SECTION 6 - UTILITAIRES INTERNES
' ============================================================

' Lance l'extracteur PDF et attend la fin.
' Priorité 1 : pdf_extractor.exe (version installée via MSI, sans Python requis)
' Priorité 2 : pdf_extractor.py via Python (mode développement)
Private Function LancerPython(cheminPDF As String, jsonSortie As String) As Integer
    Dim scriptDir As String
    scriptDir = ThisWorkbook.Path

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim cmd As String

    ' -- Priorité 1 : exe autonome --
    Dim exePath As String
    exePath = scriptDir & "\pdf_extractor.exe"
    If FichierExiste(exePath) Then
        cmd = """" & exePath & """ """ & cheminPDF & """ """ & jsonSortie & """"
        LancerPython = wsh.Run("cmd /c " & cmd, 0, True)
        Exit Function
    End If

    ' -- Priorité 2 : script Python (mode développement) --
    Dim python As String
    python = TrouverPython()
    If python = "" Then
        MsgBox "Extracteur PDF introuvable." & vbCrLf & vbCrLf & _
               "Solution 1 : Installez EasyCompta via l'installateur .msi" & vbCrLf & _
               "Solution 2 : Installez Python 3.10+ et cochez « Add Python to PATH ».", _
               vbCritical, "Extracteur manquant"
        LancerPython = 1
        Exit Function
    End If

    cmd = """" & python & """ """ & scriptDir & "\" & SCRIPT_PDF & """ " & _
          """" & cheminPDF & """ """ & jsonSortie & """"
    LancerPython = wsh.Run("cmd /c " & cmd, 0, True)
End Function

Private Function TrouverPython() As String
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
        If FichierExiste(CStr(c)) Or TestCommande(CStr(c)) Then
            TrouverPython = CStr(c)
            Exit Function
        End If
    Next c
    TrouverPython = ""
End Function

Private Function TestCommande(cmd As String) As Boolean
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    TestCommande = (wsh.Run("cmd /c " & cmd & " --version >nul 2>&1", 0, True) = 0)
    On Error GoTo 0
End Function

Private Function FichierExiste(chemin As String) As Boolean
    FichierExiste = (Len(Dir(chemin)) > 0)
End Function

' Lecture d'un fichier texte UTF-8 (JSON produit par pdf_extractor.py)
Private Function LireTexte(chemin As String) As String
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Charset = "utf-8"
    ado.Open
    ado.LoadFromFile chemin
    LireTexte = ado.ReadText
    ado.Close
    Set ado = Nothing
End Function

' Extrait la valeur d'une clé dans un JSON simple (pas de tableaux imbriqués)
Private Function JSONVal(json As String, cle As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")

    ' Cherche "cle": "valeur"
    re.Pattern = """" & cle & """\s*:\s*""([^""]*)"""
    re.IgnoreCase = True
    If re.Test(json) Then
        JSONVal = re.Execute(json)(0).SubMatches(0)
        Exit Function
    End If

    ' Cherche "cle": valeur_numerique_ou_booleen
    re.Pattern = """" & cle & """\s*:\s*([^,\}\[\n]+)"
    If re.Test(json) Then
        JSONVal = Trim(re.Execute(json)(0).SubMatches(0))
        Exit Function
    End If

    JSONVal = ""
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
Private Function ConvertirDate(valeur As String) As Variant
    On Error Resume Next
    valeur = Trim(valeur)
    If valeur = "" Then ConvertirDate = Empty : Exit Function
    valeur = Replace(valeur, "-", "/")
    valeur = Replace(valeur, ".", "/")
    Dim parties() As String
    parties = Split(valeur, "/")
    If UBound(parties) = 2 Then
        ConvertirDate = DateSerial(CInt(parties(2)), CInt(parties(1)), CInt(parties(0)))
    Else
        ConvertirDate = CDate(valeur)
    End If
    If Err.Number <> 0 Then ConvertirDate = Empty
    On Error GoTo 0
End Function

' Génère un ID unique : F2025-0001 ou C2025-0001
Private Function GenererID(ws As Worksheet, prefixe As String) As String
    Dim der As Long
    der = DerniereLigne(ws)
    Dim annee As String
    annee = CStr(Year(Now))
    If der <= 2 Then
        GenererID = prefixe & annee & "-0001"
    Else
        GenererID = prefixe & annee & "-" & Format(der, "0000")
    End If
End Function

' Retourne le numéro de la dernière ligne avec données en colonne A
Private Function DerniereLigne(ws As Worksheet) As Long
    DerniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
End Function

' Affiche un message de statut coloré dans la zone statut de PDF_Import
Private Sub Statut(ws As Worksheet, message As String, niveau As String)
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

Private Function LicGetSecret() As String
    Dim p1 As String, p2 As String, p3 As String, p4 As String, p5 As String
    p1 = "ECPT$" : p2 = "C0mpt" : p3 = "4bl3#" : p4 = "Pr0j3" : p5 = "t9!"
    LicGetSecret = p1 & p2 & p3 & p4 & p5
End Function

Public Sub VerifierOuDemanderLicence()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Dim email As String, cle As String
    On Error Resume Next
    email = wsh.RegRead(REG_PATH & "\Email")
    cle   = wsh.RegRead(REG_PATH & "\Cle")
    On Error GoTo 0

    If email <> "" And cle <> "" And LicVerifier(cle, email) Then
        Exit Sub
    End If

    ' Incrémenter le compteur d'ouvertures
    Dim count As Long
    On Error Resume Next
    count = CLng(wsh.RegRead(REG_PATH & "\OpenCount"))
    On Error GoTo 0
    count = count + 1
    wsh.RegWrite REG_PATH & "\OpenCount", count, "REG_DWORD"

    ' Popup don toutes les 5 ouvertures (1re, 6e, 11e...)
    If count Mod 5 <> 1 Then Exit Sub

    Dim rep As Integer
    rep = MsgBox("Merci d'utiliser EasyCompta !" & vbCrLf & vbCrLf & _
                 "Ce logiciel est libre et gratuit." & vbCrLf & _
                 "Si vous souhaitez soutenir le projet," & vbCrLf & _
                 "un don PayPal vous permet d'obtenir une cle d'activation." & vbCrLf & vbCrLf & _
                 "Voulez-vous faire un don maintenant ?", _
                 vbQuestion + vbYesNo, "Soutenir EasyCompta")

    If rep = vbYes Then
        wsh.Run "cmd /c start https://paypal.me/NLaurent878", 0, False
    End If

    rep = MsgBox("Avez-vous deja une cle de licence ?", _
                 vbQuestion + vbYesNo, "Activation EasyCompta")
    If rep = vbNo Then Exit Sub

    Do
        email = InputBox("Entrez votre adresse email :", "Activation", email)
        If email = "" Then Exit Sub

        cle = InputBox("Entrez votre cle de licence :" & vbCrLf & _
                       "(format : ECPT-AAAA-XXXX-XXXX)", "Activation", cle)
        If cle = "" Then Exit Sub

        If LicVerifier(cle, email) Then
            wsh.RegWrite REG_PATH & "\Email", LCase(Trim(email)), "REG_SZ"
            wsh.RegWrite REG_PATH & "\Cle",   UCase(Trim(cle)),   "REG_SZ"
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

Private Function LicVerifier(cle As String, email As String) As Boolean
    cle = UCase(Trim(cle))
    If Len(cle) <> 19 Then Exit Function
    If Left$(cle, 5) <> "ECPT-" Then Exit Function

    Dim hashCle As String
    hashCle = Mid$(cle, 11, 4) & Mid$(cle, 16, 4)

    Dim expected As String
    expected = LicHashCle(LCase(Trim(email)) & LicGetSecret())

    LicVerifier = (hashCle = Left$(expected, 8))
End Function

Private Function LicHashCle(texte As String) As String
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
    LicHashCle = Right$("00000000" & Hex$(CLng(h1) Xor CLng(h2)), 8)
End Function
