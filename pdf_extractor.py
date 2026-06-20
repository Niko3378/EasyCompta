#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extracteur automatique de données de factures PDF.
Usage : python pdf_extractor.py <chemin_pdf> <fichier_sortie.json>
"""

import sys
import json
import re
import os
import logging
from datetime import datetime
from pathlib import Path

logging.basicConfig(level=logging.WARNING, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# Détection du moteur PDF disponible
try:
    import pdfplumber
    PDF_ENGINE = 'pdfplumber'
except ImportError:
    try:
        from pypdf import PdfReader as _PdfReader
        PDF_ENGINE = 'pypdf'
    except ImportError:
        try:
            from PyPDF2 import PdfReader as _PdfReader
            PDF_ENGINE = 'PyPDF2'
        except ImportError:
            PDF_ENGINE = None


class ExtractionError(Exception):
    pass


class FactureExtractor:
    """Extrait les données comptables d'une facture PDF."""

    PATTERNS = {
        'numero_facture': [
            # "Facture N : FA-2025-0042" ou "Facture N° : FACT-2025-0042"
            r'[Ff]acture\s+N[°o]?\s*[:\-]\s*([A-Z0-9][A-Z0-9\-/_.]{2,28})',
            # "Invoice No: INV-0042"
            r'[Ii]nvoice\s+N[°oa]?\s*[:\-]\s*([A-Z0-9][A-Z0-9\-/_.]{2,28})',
            # "N° de facture : ..."
            r'N[°o]\s*de\s+[Ff]acture\s*[:\-]\s*([A-Z0-9][A-Z0-9\-/_.]{2,28})',
            # "Référence : FA-001"
            r'(?:R[ée]f[ée]rence|Ref\.?)\s*[:\-]\s*([A-Z0-9][A-Z0-9\-/_.]{4,28})',
            # "FACT-2025-0042" ou "FA-2025-001"
            r'\b((?:FACT|FAC|FA|INV|DEVIS)[-_]\d{4,6}[-_]\d{1,6})\b',
            r'\b((?:FACT|FAC|FA|INV)[-_]\d{4,10})\b',
        ],
        'date_emission': [
            # "Date : 15/05/2025" ou "Date d'émission : ..."
            r"(?:Date\s*(?:d['’])[Éé]mission|Date\s+d['']?[Éé]mission|Date\s+de\s+[Ff]acture|Date)\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
            r"(?:Date\s+d['’]?[Éé]mission|Date\s+facture)\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})",
            r'(?:Le|En\s+date\s+du?)\s*[:\-]\s*(\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{2,4})',
            # fallback : première date complète sur 4 chiffres
            r'(\d{2}[\/\-\.]\d{2}[\/\-\.]\d{4})',
        ],
        'montant_ht': [
            # "Total H.T. : 2 900,00 EUR"  — capture commence forcément par un chiffre
            r'Total\s+H\.?T\.?\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'H\.?T\.?\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'[Hh]ors[\s\-]?[Tt]ax[e]?[s]?\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'[Ss]ous[\s\-]?total\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'[Bb]ase\s+(?:de\s+calcul\s+)?T\.?V\.?A\.?\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
        ],
        'montant_tva': [
            # "TVA 20 % : 580,00 EUR"
            r'TVA\s+\d+[\s,.]?\d*\s*%\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            # "T.V.A. : 580,00"
            r'T\.?V\.?A\.?\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'TVA\s+(?:à|au|de)\s+\d+[\s,.]?\d*\s*%\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
        ],
        'montant_ttc': [
            # "NET A PAYER : 3 480,00 EUR"
            r'NET\s+A\s+PAYER\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'Net\s+[àa]\s+payer\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'TOTAL\s+TTC\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'Total\s+T\.?T\.?C\.?\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'Montant\s+total\s*[:\-]?\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
            r'[Tt]outes?\s+[Tt]axes?\s+[Cc]ompris[e]?s?\s*[:\-]\s*(\d[\d\s]{0,8}[,.]?\d{0,2})\s*(?:EUR|€)?',
        ],
        'taux_tva': [
            r'TVA\s+(\d+(?:[,.]\d+)?)\s*%',
            r'T\.?V\.?A\.?\s*[àa]?\s*(\d+(?:[,.]\d+)?)\s*%',
            r'(\d+(?:[,.]\d+)?)\s*%\s*(?:de\s+)?T\.?V\.?A',
            r'[Tt]aux\s*[:\-]\s*(\d+(?:[,.]\d+)?)\s*%',
        ],
        'fournisseur_nom': [
            r'([A-ZÉÀÈÊÎÏÔÙÛÜ][A-Za-zÀ-ÿ\s&\-]{2,40})\s+(?:SARL|SAS|SA|SASU|EURL)\b',
            r'(?:SARL|SAS|SA|SASU|EURL)\s+([A-ZÉÀÈÊÎÏÔÙÛÜ][A-Za-zÀ-ÿ\s&\-]{2,40})',
        ],
    }

    TAUX_TVA_FR = [0, 5.5, 10.0, 20.0]

    KEYWORDS_RECU = ['à payer', 'net à payer', 'règlement', 'virement', 'RIB', 'IBAN',
                     'notre référence', 'votre commande', 'bon de commande',
                     'prestataire', 'fournisseur', 'vendeur', 'émetteur']
    KEYWORDS_EMIS = ['votre référence client', 'numéro client', 'code client',
                     'facturé à', 'destinataire', 'client', 'acheteur']

    def __init__(self):
        if PDF_ENGINE is None:
            raise ExtractionError(
                "Aucune bibliothèque PDF disponible.\n"
                "Installez-en une via : pip install pdfplumber"
            )

    # ------------------------------------------------------------------
    # Extraction texte
    # ------------------------------------------------------------------

    def extraire_texte(self, chemin_pdf: str) -> str:
        p = Path(chemin_pdf)
        if not p.exists():
            raise ExtractionError(f"Fichier introuvable : {chemin_pdf}")
        if p.suffix.lower() != '.pdf':
            raise ExtractionError(f"Ce fichier n'est pas un PDF : {p.name}")

        texte = ""
        try:
            if PDF_ENGINE == 'pdfplumber':
                with pdfplumber.open(str(p)) as pdf:
                    for page in pdf.pages:
                        t = page.extract_text()
                        if t:
                            texte += t + "\n"
            else:
                reader = _PdfReader(str(p))
                for page in reader.pages:
                    t = page.extract_text()
                    if t:
                        texte += t + "\n"
        except Exception as exc:
            raise ExtractionError(f"Erreur lecture PDF : {exc}")

        if not texte.strip():
            raise ExtractionError(
                "PDF vide ou scanné (image). "
                "L'extraction automatique nécessite un PDF texte."
            )
        return texte

    # ------------------------------------------------------------------
    # Utilitaires
    # ------------------------------------------------------------------

    def _nettoyer_montant(self, valeur: str) -> float:
        """Convertit une chaîne montant en float (gère , et . comme séparateurs)."""
        if not valeur:
            return 0.0
        valeur = re.sub(r'[€$£\s ]', '', valeur).strip()
        # Format "1.234,56" → "1234.56"
        if ',' in valeur and '.' in valeur:
            if valeur.rindex('.') < valeur.rindex(','):
                valeur = valeur.replace('.', '').replace(',', '.')
            else:
                valeur = valeur.replace(',', '')
        elif ',' in valeur:
            valeur = valeur.replace(',', '.')
        try:
            result = round(float(valeur), 2)
            return result if result >= 0 else 0.0
        except ValueError:
            return 0.0

    def _normaliser_date(self, date_str: str) -> str:
        """Retourne la date au format JJ/MM/AAAA."""
        date_str = re.sub(r'[\-\.]', '/', date_str.strip())
        for fmt in ['%d/%m/%Y', '%d/%m/%y', '%Y/%m/%d']:
            try:
                return datetime.strptime(date_str, fmt).strftime('%d/%m/%Y')
            except ValueError:
                continue
        return date_str

    def _extraire_champ(self, texte: str, patterns: list) -> str:
        for pat in patterns:
            m = re.search(pat, texte, re.IGNORECASE | re.MULTILINE)
            if m:
                return m.group(1).strip()
        return ""

    def _arrondir_taux(self, taux: float) -> float:
        """Arrondit au taux TVA français le plus proche."""
        for t in self.TAUX_TVA_FR:
            if abs(taux - t) < 1.5:
                return t
        return round(taux, 1)

    def _detecter_type(self, texte: str) -> str:
        texte_lower = texte.lower()
        score_recu = sum(1 for kw in self.KEYWORDS_RECU if kw in texte_lower)
        score_emis = sum(1 for kw in self.KEYWORDS_EMIS if kw in texte_lower)
        return "Client" if score_emis > score_recu else "Fournisseur"

    # ------------------------------------------------------------------
    # Extraction principale
    # ------------------------------------------------------------------

    def extraire(self, chemin_pdf: str) -> dict:
        texte = self.extraire_texte(chemin_pdf)

        num_facture = self._extraire_champ(texte, self.PATTERNS['numero_facture'])
        date_raw    = self._extraire_champ(texte, self.PATTERNS['date_emission'])
        ht_raw      = self._extraire_champ(texte, self.PATTERNS['montant_ht'])
        tva_raw     = self._extraire_champ(texte, self.PATTERNS['montant_tva'])
        ttc_raw     = self._extraire_champ(texte, self.PATTERNS['montant_ttc'])
        taux_raw    = self._extraire_champ(texte, self.PATTERNS['taux_tva'])

        ht   = self._nettoyer_montant(ht_raw)
        tva  = self._nettoyer_montant(tva_raw)
        ttc  = self._nettoyer_montant(ttc_raw)
        taux = self._nettoyer_montant(taux_raw)

        # Reconstruction croisée si données partielles
        if ttc > 0 and ht > 0 and tva == 0:
            tva = round(ttc - ht, 2)
        elif ttc > 0 and tva > 0 and ht == 0:
            ht = round(ttc - tva, 2)
        elif ht > 0 and tva > 0 and ttc == 0:
            ttc = round(ht + tva, 2)

        # Calcul taux si absent
        if taux == 0 and ht > 0 and tva > 0:
            taux = self._arrondir_taux(round((tva / ht) * 100, 1))
        elif taux > 0:
            taux = self._arrondir_taux(taux)

        date = self._normaliser_date(date_raw) if date_raw else ""
        type_facture = self._detecter_type(texte)

        avertissements = []
        if not num_facture:
            avertissements.append("Numero de facture non detecte")
        if not date:
            avertissements.append("Date d'emission non detectee")
        if ttc == 0:
            avertissements.append("Montant TTC non detecte - verification manuelle requise")

        return {
            'succes': True,
            'numero_facture': num_facture,
            'date_emission': date,
            'montant_ht': ht,
            'montant_tva': tva,
            'montant_ttc': ttc,
            'taux_tva': taux,
            'type_facture': type_facture,
            'chemin_pdf': str(Path(chemin_pdf).resolve()),
            'date_extraction': datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
            'moteur_pdf': PDF_ENGINE,
            'avertissements': ' | '.join(avertissements),
        }


# ------------------------------------------------------------------
# Point d'entrée CLI (appelé par VBA via Shell)
# ------------------------------------------------------------------

def main():
    if len(sys.argv) < 3:
        sortie = {
            'succes': False,
            'erreur': 'Usage : python pdf_extractor.py <chemin_pdf> <sortie.json>',
        }
        print(json.dumps(sortie, ensure_ascii=False))
        sys.exit(1)

    chemin_pdf   = sys.argv[1]
    fichier_json = sys.argv[2]

    try:
        extractor = FactureExtractor()
        donnees   = extractor.extraire(chemin_pdf)
    except ExtractionError as exc:
        donnees = {'succes': False, 'erreur': str(exc), 'type_facture': ''}
    except Exception as exc:
        donnees = {'succes': False, 'erreur': f"Erreur inattendue : {exc}", 'type_facture': ''}

    os.makedirs(os.path.dirname(os.path.abspath(fichier_json)), exist_ok=True)
    with open(fichier_json, 'w', encoding='cp1252') as f:
        json.dump(donnees, f, ensure_ascii=False, indent=2)

    # Signal de fin lisible par VBA
    print("EXTRACTION_DONE")
    sys.exit(0)


if __name__ == '__main__':
    main()
