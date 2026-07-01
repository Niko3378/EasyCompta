#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Générateur de clés de licence EasyCompta.
USAGE PRIVÉ — ne pas distribuer, ne pas inclure dans le MSI.

Modes d'utilisation :
    # Clé unique
    python generer_cle.py email@exemple.com
    python generer_cle.py email@exemple.com 2026

    # Lot — plusieurs emails en arguments
    python generer_cle.py email1@a.com email2@b.com email3@c.com

    # Lot — depuis un fichier texte (un email par ligne)
    python generer_cle.py --lot emails.txt
    python generer_cle.py --lot emails.txt --annee 2026

    # Export CSV
    python generer_cle.py --lot emails.txt --csv cles.csv
"""

import sys
import csv
import datetime
from pathlib import Path

# Doit être identique à la constante SECRET dans vba_Comptabilite.bas
SECRET = "ECPT$C0mpt4bl3#Pr0j3t9!"


def _hash_cle(texte: str) -> str:
    """Même algorithme que LicHashCle() dans vba_Comptabilite.bas."""
    h1 = 1505843
    h2 = 1234567891
    for c in texte:
        v = ord(c)
        h1 = ((h1 & 0x3FFFFFF) * 31 + v) & 0x7FFFFFFF
        h2 = ((h2 & 0x3FFFFFF) * 37 + v) & 0x7FFFFFFF
    return format(h1 ^ h2, '08X')


def generer_cle(email: str, annee: int | None = None) -> str:
    if annee is None:
        annee = datetime.date.today().year
    h = _hash_cle(email.lower().strip() + SECRET)
    return f"ECPT-{annee}-{h[:4]}-{h[4:8]}"


def verifier_cle(cle: str, email: str) -> bool:
    cle = cle.upper().strip()
    if len(cle) != 19:
        return False
    if not cle.startswith("ECPT-"):
        return False
    hash_cle = cle[10:14] + cle[15:19]
    expected = _hash_cle(email.lower().strip() + SECRET)
    return hash_cle == expected[:8]


def _lire_emails(chemin: str) -> list[str]:
    emails = []
    for ligne in Path(chemin).read_text(encoding="utf-8").splitlines():
        ligne = ligne.strip()
        if ligne and not ligne.startswith("#"):
            emails.append(ligne)
    return emails


def _afficher_lot(emails: list[str], annee: int | None):
    print()
    print(f"  {'EMAIL':<35} {'CLÉ':<20} VÉRIF")
    print(f"  {'-'*35} {'-'*20} -----")
    for email in emails:
        cle = generer_cle(email, annee)
        ok = "OK" if verifier_cle(cle, email) else "ERREUR"
        print(f"  {email:<35} {cle:<20} {ok}")
    print()


def _exporter_csv(emails: list[str], annee: int | None, chemin_csv: str):
    with open(chemin_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["Email", "Clé", "Année"])
        for email in emails:
            cle = generer_cle(email, annee)
            w.writerow([email, cle, annee or datetime.date.today().year])
    print(f"  Export CSV : {chemin_csv} ({len(emails)} clés)")


if __name__ == "__main__":
    args = sys.argv[1:]

    if not args:
        print(__doc__)
        sys.exit(1)

    annee = None
    csv_out = None

    # Extraire options --annee et --csv
    filtered = []
    i = 0
    while i < len(args):
        if args[i] == "--annee" and i + 1 < len(args):
            annee = int(args[i + 1]); i += 2
        elif args[i] == "--csv" and i + 1 < len(args):
            csv_out = args[i + 1]; i += 2
        else:
            filtered.append(args[i]); i += 1
    args = filtered

    # Mode lot depuis fichier
    if args and args[0] == "--lot":
        if len(args) < 2:
            print("Usage: python generer_cle.py --lot emails.txt")
            sys.exit(1)
        emails = _lire_emails(args[1])
        if not emails:
            print("Aucun email trouvé dans le fichier.")
            sys.exit(1)
        _afficher_lot(emails, annee)
        if csv_out:
            _exporter_csv(emails, annee, csv_out)

    # Plusieurs emails en arguments
    elif len(args) > 1 or (len(args) == 1 and "@" in args[0]):
        emails = [a for a in args if "@" in a]
        if not emails:
            print("Aucun email valide fourni.")
            sys.exit(1)
        if len(emails) == 1:
            email = emails[0]
            cle = generer_cle(email, annee)
            print()
            print(f"  Email : {email}")
            print(f"  Clé   : {cle}")
            print(f"  Vérif : {'OK' if verifier_cle(cle, email) else 'ERREUR'}")
            print()
            print("  -> Envoyez cette clé au client par email.")
            print()
        else:
            _afficher_lot(emails, annee)
            if csv_out:
                _exporter_csv(emails, annee, csv_out)

    else:
        print(f"Argument non reconnu : {args}")
        sys.exit(1)
