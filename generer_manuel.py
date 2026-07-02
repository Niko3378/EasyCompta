#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère le manuel utilisateur PDF de EasyCompta (fpdf2).
Sortie : Manuel_EasyCompta.pdf
"""

import os
import sys
from fpdf import FPDF
from fpdf.enums import XPos, YPos

# Polices TrueType Windows (support Unicode complet pour les accents français)
_FONTS = {
    '':   r"C:\Windows\Fonts\arial.ttf",
    'B':  r"C:\Windows\Fonts\arialbd.ttf",
    'I':  r"C:\Windows\Fonts\ariali.ttf",
    'BI': r"C:\Windows\Fonts\arialbi.ttf",
}
FONT = 'Arial'

BASE    = os.path.dirname(os.path.abspath(__file__))
SORTIE  = os.path.join(BASE, "Manuel_EasyCompta.pdf")
with open(os.path.join(BASE, "version.txt")) as _vf:
    VERSION_APP = _vf.read().strip()

# ── Palette ──────────────────────────────────────────────────────────────────
C_BLEU    = (31,  78, 121)
C_BLEU2   = (46, 117, 182)
C_VERT    = (55, 86,  35)
C_ORANGE  = (197, 90, 17)
C_ROUGE   = (192,  0,  0)
C_VIOLET  = (112, 48, 160)
C_GRIS    = (242, 242, 242)
C_TEXTE   = (30,  30,  30)
C_BLANC   = (255, 255, 255)


class ManuelPDF(FPDF):
    def __init__(self):
        super().__init__('P', 'mm', 'A4')
        self.set_auto_page_break(auto=True, margin=20)
        self.set_margins(20, 20, 20)
        for style, path in _FONTS.items():
            self.add_font(FONT, style=style, fname=path)

    # ── Header / Footer ──────────────────────────────────────────────────────

    def header(self):
        if self.page_no() == 1:
            return
        self.set_fill_color(*C_BLEU)
        self.rect(0, 0, 210, 10, 'F')
        self.set_font(FONT, 'B', 8)
        self.set_text_color(*C_BLANC)
        self.set_xy(10, 2)
        self.cell(0, 6, 'EasyCompta — Manuel Utilisateur', align='L')
        self.set_xy(10, 2)
        self.cell(0, 6, f'Page {self.page_no()}', align='R')
        self.set_text_color(*C_TEXTE)
        self.ln(8)

    def footer(self):
        if self.page_no() == 1:
            return
        self.set_y(-15)
        self.set_font(FONT, 'I', 7)
        self.set_text_color(150, 150, 150)
        self.cell(0, 6, 'EasyCompta — Logiciel gratuit — https://paypal.me/NLaurent878', align='C')
        self.set_text_color(*C_TEXTE)

    # ── Helpers de mise en forme ─────────────────────────────────────────────

    def titre_section(self, numero, texte, couleur=C_BLEU):
        self.ln(4)
        self.set_fill_color(*couleur)
        self.set_text_color(*C_BLANC)
        self.set_font(FONT, 'B', 13)
        self.cell(0, 9, f'  {numero}  {texte}', fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(*C_TEXTE)
        self.ln(3)

    def sous_titre(self, texte, couleur=C_BLEU2):
        self.ln(2)
        self.set_fill_color(*couleur)
        self.set_text_color(*C_BLANC)
        self.set_font(FONT, 'B', 10)
        self.cell(0, 7, f'   {texte}', fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_text_color(*C_TEXTE)
        self.ln(2)

    def paragraphe(self, texte):
        self.set_font(FONT, '', 10)
        self.set_text_color(*C_TEXTE)
        self.multi_cell(0, 6, texte, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.ln(1)

    def etape(self, numero, texte):
        self.set_font(FONT, 'B', 10)
        self.set_fill_color(*C_BLEU2)
        self.set_text_color(*C_BLANC)
        self.cell(7, 6, str(numero), fill=True, align='C')
        self.set_font(FONT, '', 10)
        self.set_text_color(*C_TEXTE)
        self.cell(0, 6, f'  {texte}', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def puce(self, texte, couleur_puce=C_BLEU2):
        self.set_font(FONT, '', 10)
        self.set_text_color(*C_TEXTE)
        self.multi_cell(0, 6, f'•  {texte}', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    def encadre(self, titre, texte, couleur=C_GRIS, couleur_titre=C_BLEU):
        self.ln(2)
        self.set_fill_color(*couleur)
        self.set_draw_color(*couleur_titre)
        self.set_line_width(0.5)
        # Titre encadré
        self.set_font(FONT, 'B', 9)
        self.set_fill_color(*couleur_titre)
        self.set_text_color(*C_BLANC)
        self.cell(0, 6, f'  {titre}', fill=True,
                  new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        # Contenu
        self.set_fill_color(*couleur)
        self.set_text_color(*C_TEXTE)
        self.set_font(FONT, '', 9)
        self.multi_cell(0, 5.5, texte, fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_draw_color(0, 0, 0)
        self.set_line_width(0.2)
        self.ln(2)

    def tableau_tva(self):
        lignes = [
            ("0 %",    "Exonération",     "Exportations, certains services médicaux"),
            ("5,5 %",  "Taux réduit",     "Alimentation, livres, abonnements énergie"),
            ("10 %",   "Taux intermédiaire", "Restauration, transport, travaux"),
            ("20 %",   "Taux normal",     "Majorité des biens et services"),
        ]
        cols = [22, 40, 108]
        hdrs = ["Taux", "Catégorie", "Exemples"]
        self.set_font(FONT, 'B', 9)
        self.set_fill_color(*C_BLEU)
        self.set_text_color(*C_BLANC)
        for w, h in zip(cols, hdrs):
            self.cell(w, 7, f'  {h}', fill=True, border=1)
        self.ln()
        self.set_font(FONT, '', 9)
        self.set_text_color(*C_TEXTE)
        for i, (taux, cat, ex) in enumerate(lignes):
            bg = C_GRIS if i % 2 == 0 else C_BLANC
            self.set_fill_color(*bg)
            self.cell(cols[0], 6, f'  {taux}',  fill=True, border=1)
            self.cell(cols[1], 6, f'  {cat}',   fill=True, border=1)
            self.cell(cols[2], 6, f'  {ex}',    fill=True, border=1)
            self.ln()
        self.ln(2)


# ── Construction du document ──────────────────────────────────────────────────

def generer():
    pdf = ManuelPDF()

    # ─────────────────────────────────────────────────────────────────────────
    # PAGE 1 — Couverture
    # ─────────────────────────────────────────────────────────────────────────
    pdf.add_page()

    # Bande titre
    pdf.set_fill_color(*C_BLEU)
    pdf.rect(0, 0, 210, 297, 'F')

    pdf.set_font(FONT, 'B', 38)
    pdf.set_text_color(*C_BLANC)
    pdf.set_xy(0, 70)
    pdf.cell(210, 20, 'EasyCompta', align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(FONT, '', 18)
    pdf.set_text_color(189, 215, 238)
    pdf.cell(210, 12, 'Système de gestion comptable', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Bandeau central blanc
    pdf.set_fill_color(*C_BLANC)
    pdf.rect(0, 130, 210, 50, 'F')

    pdf.set_font(FONT, 'B', 14)
    pdf.set_text_color(*C_BLEU)
    pdf.set_xy(0, 140)
    pdf.cell(210, 10, 'Manuel Utilisateur', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_font(FONT, '', 11)
    pdf.set_text_color(*C_TEXTE)
    pdf.cell(210, 8, 'Import PDF automatique  •  TVA mensuelle  •  Tableaux de bord',
             align='C', new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Version et date
    pdf.set_font(FONT, 'I', 9)
    pdf.set_text_color(150, 150, 150)
    pdf.set_xy(0, 220)
    pdf.cell(210, 7, f'Version {VERSION_APP}  —  2026', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # PayPal mention
    pdf.set_fill_color(*C_ORANGE)
    pdf.rect(40, 250, 130, 20, 'F')
    pdf.set_font(FONT, 'B', 10)
    pdf.set_text_color(*C_BLANC)
    pdf.set_xy(40, 254)
    pdf.cell(130, 6, 'Logiciel gratuit — Soutenez-nous !', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(FONT, '', 9)
    pdf.set_xy(40, 260)
    pdf.cell(130, 6, 'paypal.me/NLaurent878', align='C')

    # ─────────────────────────────────────────────────────────────────────────
    # PAGE 2 — Sommaire
    # ─────────────────────────────────────────────────────────────────────────
    pdf.add_page()
    pdf.set_text_color(*C_TEXTE)

    pdf.set_font(FONT, 'B', 16)
    pdf.set_fill_color(*C_BLEU)
    pdf.set_text_color(*C_BLANC)
    pdf.cell(0, 11, '  Table des matières', fill=True,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*C_TEXTE)
    pdf.ln(4)

    chapitres = [
        ("1",  "Présentation du logiciel",              "3"),
        ("2",  "Installation",                           "3"),
        ("3",  "Importer une facture PDF",               "4"),
        ("4",  "Saisie manuelle d'une facture",          "5"),
        ("5",  "Base de données Fournisseurs",            "5"),
        ("6",  "Base de données Clients",                 "6"),
        ("7",  "Gestion du statut de paiement",          "6"),
        ("8",  "Export CSV comptable",                   "7"),
        ("9",  "Calcul de TVA mensuelle",                "7"),
        ("10", "Tableaux de bord analytiques",           "8"),
        ("11", "Taux de TVA français",                   "9"),
        ("12", "Limites et conseils",                    "9"),
        ("13", "Soutenir le projet",                     "10"),
    ]
    for num, titre, page in chapitres:
        pdf.set_font(FONT, 'B', 10)
        pdf.set_text_color(*C_BLEU)
        pdf.cell(10, 7, num)
        pdf.set_font(FONT, '', 10)
        pdf.set_text_color(*C_TEXTE)
        pdf.cell(145, 7, titre)
        pdf.set_font(FONT, 'I', 9)
        pdf.set_text_color(120, 120, 120)
        pdf.cell(0, 7, page, align='R', new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        # Ligne pointillée
        pdf.set_draw_color(200, 200, 200)
        pdf.set_line_width(0.1)
        pdf.line(20, pdf.get_y(), 190, pdf.get_y())

    # ─────────────────────────────────────────────────────────────────────────
    # PAGE 3+ — Contenu
    # ─────────────────────────────────────────────────────────────────────────
    pdf.add_page()

    # ── Section 1 : Présentation ─────────────────────────────────────────────
    pdf.titre_section("1", "Présentation du logiciel")
    pdf.paragraphe(
        "EasyCompta est un logiciel de gestion comptable autonome, conçu pour les "
        "professionnels et TPE souhaitant gérer leurs factures fournisseurs et clients "
        "sans abonnement cloud ni installation complexe.\n\n"
        "Il repose sur Microsoft Excel pour l'interface et stocke toutes les données "
        "localement sur votre ordinateur. Un extracteur PDF intégré analyse automatiquement "
        "vos factures et pré-remplit les champs comptables."
    )

    pdf.sous_titre("Fonctionnalités principales")
    for f in [
        "Import et extraction automatique de données depuis des factures PDF",
        "Base de données Fournisseurs et Clients (jusqu'à 2000 entrées chacune)",
        "Gestion du statut de paiement avec colorisation automatique des lignes",
        "Export CSV comptable (compatible Excel, Sage, EBP, Cegid)",
        "Calcul automatique de la TVA mensuelle et du solde cumulé",
        "Tableaux de bord avec graphiques (CA, dépenses, TVA)",
        "Fonctionne 100 % hors ligne — aucune connexion requise",
    ]:
        pdf.puce(f)
    pdf.ln(2)

    # ── Section 2 : Installation ─────────────────────────────────────────────
    pdf.titre_section("2", "Installation")
    pdf.paragraphe(
        "L'installation se fait via un fichier .msi standard Windows. "
        "Aucune connaissance technique n'est requise."
    )
    pdf.sous_titre("Étapes d'installation")
    for i, e in enumerate([
        "Double-cliquez sur EasyCompta_v2.0.0.msi",
        "Acceptez les conditions d'utilisation",
        "Choisissez le dossier d'installation (ou laissez le dossier proposé)",
        "Cliquez sur Installer",
        "Une fois terminé, un raccourci EasyCompta apparaît sur votre Bureau",
    ], 1):
        pdf.etape(i, e)
    pdf.ln(2)

    pdf.encadre(
        "Prérequis",
        "  - Microsoft Excel 2016 ou version ultérieure (obligatoire)\n"
        "  - Windows 10 ou Windows 11 (64 bits)\n"
        "  - Python non requis (l'extracteur PDF est intégré au logiciel)",
        couleur=C_GRIS, couleur_titre=C_VERT
    )

    pdf.sous_titre("Première ouverture")
    for i, e in enumerate([
        "Double-cliquez sur le raccourci EasyCompta depuis votre Bureau",
        "Excel s'ouvre — si demandé, cliquez sur Activer les macros",
        "Le logiciel est prêt à l'emploi",
    ], 1):
        pdf.etape(i, e)

    # ── Section 3 : Import PDF ───────────────────────────────────────────────
    pdf.add_page()
    pdf.titre_section("3", "Importer une facture PDF", C_ORANGE)
    pdf.paragraphe(
        "Cette fonctionnalité analyse automatiquement un fichier PDF de facture "
        "et en extrait les données comptables clés : numéro, date, montants HT/TVA/TTC "
        "et type (Fournisseur ou Client)."
    )

    pdf.sous_titre("Procédure d'import")
    for i, e in enumerate([
        "Ouvrez la feuille PDF_Import (onglet orange)",
        "Cliquez sur le bouton Importer PDF",
        "Dans la fenêtre qui s'ouvre, naviguez jusqu'à votre facture et cliquez Ouvrir",
        "Cliquez sur Extraire données — patientez quelques secondes",
        "Vérifiez les champs extraits et corrigez si nécessaire",
        "Cliquez sur Envoyer vers base pour enregistrer la facture",
    ], 1):
        pdf.etape(i, e)
    pdf.ln(2)

    pdf.encadre(
        "Détection automatique du type de facture",
        "  Le logiciel analyse le contenu du PDF pour déterminer s'il s'agit :\n"
        "  - d'une facture FOURNISSEUR (que vous avez reçue et que vous payez)\n"
        "  - d'une facture CLIENT (que vous avez émise et que vous encaissez)\n\n"
        "  Si la détection est incorrecte, corrigez manuellement le champ Type de facture "
        "avant d'envoyer vers la base.",
        couleur_titre=C_ORANGE
    )

    pdf.sous_titre("Données extraites automatiquement")
    champs = [
        ("Numéro de facture", "Référence unique de la facture"),
        ("Date d'émission",   "Date de la facture (format JJ/MM/AAAA)"),
        ("Montant HT",        "Montant hors taxes"),
        ("Montant TVA",       "Montant de la taxe"),
        ("Montant TTC",       "Total toutes taxes comprises"),
        ("Taux TVA",          "Taux appliqué (0 / 5,5 / 10 / 20 %)"),
        ("Type de facture",   "Fournisseur ou Client"),
    ]
    pdf.set_font(FONT, 'B', 9)
    pdf.set_fill_color(*C_BLEU)
    pdf.set_text_color(*C_BLANC)
    pdf.cell(55, 6, '  Champ', fill=True, border=1)
    pdf.cell(0,  6, '  Description', fill=True, border=1,
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(FONT, '', 9)
    pdf.set_text_color(*C_TEXTE)
    for i, (champ, desc) in enumerate(champs):
        bg = C_GRIS if i % 2 == 0 else C_BLANC
        pdf.set_fill_color(*bg)
        pdf.cell(55, 5.5, f'  {champ}', fill=True, border=1)
        pdf.cell(0,  5.5, f'  {desc}',  fill=True, border=1,
                 new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.ln(3)

    # ── Section 4 : Saisie manuelle ──────────────────────────────────────────
    pdf.titre_section("4", "Saisie manuelle d'une facture")
    pdf.paragraphe(
        "Vous pouvez saisir des factures directement dans les feuilles de base de données, "
        "sans passer par l'import PDF."
    )
    for i, e in enumerate([
        "Ouvrez la feuille Fournisseurs_DB ou Clients_DB",
        "Descendez à la première ligne vide sous les données existantes",
        "Laissez la colonne ID_Facture vide (générée automatiquement par la macro)",
        "Remplissez les colonnes : nom, numéro, date, montants, taux TVA",
        "Utilisez les listes déroulantes pour le Statut_Paiement et le Taux_TVA",
    ], 1):
        pdf.etape(i, e)

    # ── Section 5 : Fournisseurs ─────────────────────────────────────────────
    pdf.add_page()
    pdf.titre_section("5", "Base de données Fournisseurs", C_BLEU)
    pdf.paragraphe(
        "La feuille Fournisseurs_DB contient toutes les factures reçues "
        "(achats, charges, prestataires). Elle peut accueillir jusqu'à 2000 entrées."
    )

    pdf.sous_titre("Colonnes de la base Fournisseurs")
    cols_fourn = [
        ("ID_Facture",       "Identifiant unique généré automatiquement (ex : F2026-0001)"),
        ("Fournisseur",      "Nom du fournisseur ou prestataire"),
        ("Numéro_Facture",   "Référence de la facture"),
        ("Date_Emission",    "Date de la facture"),
        ("Montant_HT",       "Montant hors taxes (€)"),
        ("Montant_TVA",      "Montant de la TVA (€)"),
        ("Montant_TTC",      "Total TTC (€)"),
        ("Taux_TVA",         "Taux de TVA appliqué (%)"),
        ("Catégorie",        "Catégorie comptable (libre)"),
        ("Lien_PDF",         "Chemin vers le fichier PDF source"),
        ("Statut_Paiement",  "Payé / En attente / En retard / Annulé"),
        ("Date_Import",      "Date d'enregistrement dans le logiciel"),
        ("Source_Extraction","Manuel ou Automatique"),
        ("Date_Règlement",   "Date effective du paiement (renseignée via le bouton Statut paiement)"),
    ]
    for col, desc in cols_fourn:
        pdf.puce(f"{col} : {desc}")
    pdf.ln(2)

    pdf.encadre(
        "Astuce — Ouvrir le PDF depuis la base",
        "  Depuis Fournisseurs_DB, cliquez sur une ligne de données puis "
        "sur le bouton Ouvrir PDF pour visualiser la facture source directement.",
        couleur_titre=C_BLEU
    )

    # ── Section 6 : Clients ──────────────────────────────────────────────────
    pdf.titre_section("6", "Base de données Clients", C_VERT)
    pdf.paragraphe(
        "La feuille Clients_DB contient toutes les factures émises "
        "(ventes, prestations facturées). Même structure que Fournisseurs_DB."
    )
    pdf.puce("Les montants des Clients contribuent à la TVA collectée dans le calcul TVA.")
    pdf.puce("Les montants des Fournisseurs contribuent à la TVA déductible.")
    pdf.puce("Le solde TVA à payer = TVA collectée (Clients) − TVA déductible (Fournisseurs).")

    # ── Section 7 : Statut paiement ──────────────────────────────────────────
    pdf.add_page()
    pdf.titre_section("7", "Gestion du statut de paiement", C_ORANGE)
    pdf.paragraphe(
        "Le bouton orange Statut paiement, disponible sur les feuilles Fournisseurs_DB "
        "et Clients_DB, permet de mettre à jour rapidement le statut de règlement d'une "
        "facture et de visualiser en un coup d'oeil l'état de votre trésorerie."
    )

    pdf.sous_titre("Procédure")
    for i, e in enumerate([
        "Ouvrez Fournisseurs_DB ou Clients_DB",
        "Cliquez sur la ligne de la facture à mettre à jour",
        "Cliquez sur le bouton orange Statut paiement",
        "Choisissez Oui pour marquer Payé, Non pour un autre statut",
        "Si Payé : saisissez la date de règlement (par défaut aujourd'hui)",
        "La ligne se colore immédiatement selon le nouveau statut",
    ], 1):
        pdf.etape(i, e)
    pdf.ln(2)

    pdf.sous_titre("Codes couleur")
    couleurs = [
        ("Vert",   "Payé",        "Facture réglée — date de règlement enregistrée"),
        ("Jaune",  "En attente",  "Facture en cours — paiement non encore reçu/effectué"),
        ("Rouge",  "En retard",   "Facture échue non réglée — relance recommandée"),
        ("Gris",   "Annulé",      "Facture annulée ou avoir émis"),
    ]
    cols_w = [20, 28, 122]
    hdrs_c = ["Couleur", "Statut", "Signification"]
    pdf.set_font(FONT, 'B', 9)
    pdf.set_fill_color(*C_BLEU)
    pdf.set_text_color(*C_BLANC)
    for w, h in zip(cols_w, hdrs_c):
        pdf.cell(w, 6, f'  {h}', fill=True, border=1)
    pdf.ln()
    pdf.set_font(FONT, '', 9)
    pdf.set_text_color(*C_TEXTE)
    for i, (coul, statut, signif) in enumerate(couleurs):
        bg = C_GRIS if i % 2 == 0 else C_BLANC
        pdf.set_fill_color(*bg)
        pdf.cell(cols_w[0], 6, f'  {coul}',  fill=True, border=1)
        pdf.cell(cols_w[1], 6, f'  {statut}', fill=True, border=1)
        pdf.cell(cols_w[2], 6, f'  {signif}', fill=True, border=1)
        pdf.ln()
    pdf.ln(3)

    pdf.encadre(
        "Colorisation automatique à l'ouverture",
        "  À chaque ouverture de EasyCompta, toutes les lignes des deux bases de données\n"
        "  sont automatiquement colorisées selon leur statut de paiement.\n"
        "  La colonne Date_Règlement (colonne 14) enregistre la date effective du paiement.",
        couleur_titre=C_ORANGE
    )

    # ── Section 8 : Export CSV ───────────────────────────────────────────────
    pdf.titre_section("8", "Export CSV comptable", C_VERT)
    pdf.paragraphe(
        "Le bouton vert Exporter CSV, situé sur la feuille TVA_Mensuelle, génère un fichier "
        "CSV contenant l'ensemble des factures fournisseurs et clients. Ce fichier est "
        "directement exploitable par votre comptable ou importable dans un logiciel de "
        "comptabilité (Sage, EBP, Cegid, etc.)."
    )

    pdf.sous_titre("Procédure d'export")
    for i, e in enumerate([
        "Ouvrez la feuille TVA_Mensuelle",
        "Cliquez sur le bouton vert Exporter CSV",
        "Confirmez l'export dans la boîte de dialogue",
        "Le fichier EasyCompta_export_AAAA-MM.csv est créé dans le dossier du logiciel",
        "Choisissez si vous souhaitez ouvrir le fichier immédiatement",
    ], 1):
        pdf.etape(i, e)
    pdf.ln(2)

    pdf.encadre(
        "Format du fichier CSV",
        "  Séparateur : point-virgule (;) — standard français\n"
        "  Encodage : UTF-8 avec BOM (ouverture directe dans Excel sans problème d'accents)\n"
        "  Colonnes : Type ; N° Facture ; Date ; Nom ; Montant HT ; Montant TVA ;\n"
        "             Montant TTC ; Taux TVA % ; Catégorie ; Statut ; Référence PDF\n"
        "  Contenu : toutes les factures Fournisseurs et Clients confondues",
        couleur_titre=C_VERT
    )

    # ── Section 9 : TVA ──────────────────────────────────────────────────────
    pdf.add_page()
    pdf.titre_section("9", "Calcul de TVA mensuelle", C_ROUGE)
    pdf.paragraphe(
        "La feuille TVA_Mensuelle calcule automatiquement la TVA à déclarer chaque mois, "
        "en agrégeant les données de Fournisseurs_DB et Clients_DB."
    )

    pdf.sous_titre("Sélectionner l'année fiscale")
    pdf.paragraphe(
        "Modifiez la cellule B3 de la feuille TVA_Mensuelle pour changer l'année affichée. "
        "Tous les calculs se mettent à jour automatiquement."
    )

    pdf.sous_titre("Lecture du tableau TVA")
    for item in [
        "TVA Collectée : montant TVA de toutes les factures Clients du mois",
        "TVA Déductible : montant TVA de toutes les factures Fournisseurs du mois",
        "TVA à Payer : TVA Collectée − TVA Déductible (versement à l'État)",
        "Solde Cumulé : cumul annuel de la TVA à payer",
        "Statut : indique si la TVA du mois est à payer ou en crédit",
        "Échéance : date limite de déclaration (19 du mois suivant, régime mensuel)",
    ]:
        pdf.puce(item)
    pdf.ln(2)

    pdf.encadre(
        "Synthèse par taux de TVA",
        "  En bas de la feuille TVA_Mensuelle, un tableau de synthèse ventile\n"
        "  la TVA par taux (0 %, 5,5 %, 10 %, 20 %) pour l'année sélectionnée.",
        couleur_titre=C_ROUGE
    )

    # ── Section 8 : Tableaux de bord ─────────────────────────────────────────
    pdf.titre_section("10", "Tableaux de bord analytiques", C_VERT)
    pdf.paragraphe(
        "La feuille Pivot_Analytics offre une vue graphique de l'activité annuelle. "
        "Elle se synchronise avec l'année choisie dans TVA_Mensuelle (cellule B3)."
    )

    pdf.sous_titre("Graphiques disponibles")
    for g in [
        "CA vs Dépenses mensuels HT : comparaison revenus / charges mois par mois",
        "TVA collectée vs TVA déductible : visualisation de l'équilibre de TVA",
        "Évolution du solde TVA : courbe cumulée sur l'année",
    ]:
        pdf.puce(g)
    pdf.ln(2)

    pdf.sous_titre("Tableau de synthèse mensuelle")
    for col in [
        "Mois : de Janvier à Décembre",
        "CA HT (Clients) : chiffre d'affaires mensuel hors taxes",
        "Dépenses HT (Fournisseurs) : charges mensuelles hors taxes",
        "TVA Collectée / Déductible / Solde : résumé TVA du mois",
    ]:
        pdf.puce(col)
    pdf.ln(2)

    pdf.encadre(
        "Rafraîchir les tableaux de bord",
        "  Après un import groupé de plusieurs factures, cliquez sur le bouton\n"
        "  Rafraichir tableaux dans la feuille Pivot_Analytics pour forcer\n"
        "  la mise à jour des formules et graphiques.",
        couleur_titre=C_VERT
    )

    # ── Section 9 : Taux TVA ─────────────────────────────────────────────────
    pdf.add_page()
    pdf.titre_section("11", "Taux de TVA français", C_ORANGE)
    pdf.paragraphe(
        "Le logiciel gère les quatre taux de TVA en vigueur en France métropolitaine."
    )
    pdf.tableau_tva()

    pdf.encadre(
        "Taux réduits spéciaux (DOM-TOM)",
        "  Des taux spécifiques s'appliquent dans les départements et régions d'outre-mer\n"
        "  (Guadeloupe, Martinique, La Réunion). Consultez le site impots.gouv.fr\n"
        "  pour les taux en vigueur dans votre département.",
        couleur_titre=C_ORANGE
    )

    # ── Section 10 : Limites ─────────────────────────────────────────────────
    pdf.titre_section("12", "Limites et conseils", C_ROUGE)

    pdf.sous_titre("Extraction PDF automatique")
    for l in [
        "PDFs scannés (images) : non pris en charge — seuls les PDFs texte sont analysés",
        "Formats très non-standard : l'extraction peut être partielle",
        "Toujours vérifier les données extraites avant envoi en base de données",
        "En cas d'erreur, les champs peuvent être saisis manuellement",
    ]:
        pdf.puce(l, couleur_puce=C_ROUGE)
    pdf.ln(2)

    pdf.sous_titre("Conseils généraux")
    for c in [
        "Effectuez une sauvegarde régulière de EasyCompta.xlsm",
        "Ne modifiez pas les noms des feuilles (Fournisseurs_DB, Clients_DB…)",
        "Ne supprimez pas les colonnes des tables structurées Excel",
        "Pour changer d'année fiscale, modifiez uniquement la cellule B3 de TVA_Mensuelle",
        "En cas de problème, relancez la macro Setup_Boutons (Alt+F8)",
    ]:
        pdf.puce(c, couleur_puce=C_BLEU2)
    pdf.ln(2)

    pdf.encadre(
        "Sauvegarde recommandée",
        "  Le fichier EasyCompta.xlsm contient toutes vos données comptables.\n"
        "  Sauvegardez-le régulièrement sur un support externe ou dans un dossier\n"
        "  cloud (OneDrive, Google Drive) pour éviter toute perte.",
        couleur_titre=C_VERT
    )

    # ── Section 11 : Soutenir ────────────────────────────────────────────────
    pdf.titre_section("13", "Soutenir le projet", C_ORANGE)
    pdf.paragraphe(
        "EasyCompta est un logiciel entièrement gratuit, développé et maintenu "
        "bénévolement. Si il vous fait gagner du temps et vous est utile au quotidien, "
        "vous pouvez soutenir son développement par un don libre."
    )
    pdf.ln(2)

    pdf.set_fill_color(*C_ORANGE)
    pdf.rect(30, pdf.get_y(), 150, 28, 'F')
    pdf.set_font(FONT, 'B', 14)
    pdf.set_text_color(*C_BLANC)
    pdf.set_xy(30, pdf.get_y() + 4)
    pdf.cell(150, 8, 'Faire un don via PayPal', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font(FONT, '', 12)
    pdf.set_xy(30, pdf.get_y() + 2)
    pdf.cell(150, 8, 'https://paypal.me/NLaurent878', align='C',
             new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_text_color(*C_TEXTE)
    pdf.ln(6)

    pdf.paragraphe(
        "Chaque contribution, même modeste, permet de continuer à améliorer "
        "le logiciel et à fournir de nouvelles fonctionnalités.\n\n"
        "Merci pour votre soutien !"
    )

    # ─────────────────────────────────────────────────────────────────────────
    # Sauvegarde
    # ─────────────────────────────────────────────────────────────────────────
    pdf.output(SORTIE)
    print(f"  Manuel généré : {SORTIE}")


if __name__ == "__main__":
    import sys
    print("\n=== Génération du manuel PDF ===")
    try:
        generer()
    except Exception as exc:
        print(f"ERREUR : {exc}", file=sys.stderr)
        sys.exit(1)
