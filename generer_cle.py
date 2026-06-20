#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Générateur de clés de licence EasyCompta.
USAGE PRIVÉ — ne pas distribuer, ne pas inclure dans le MSI.

Usage :
    python generer_cle.py email@exemple.com
    python generer_cle.py email@exemple.com 2026
"""

import hashlib
import sys
import datetime

# Doit être identique à la constante SECRET dans vba_Comptabilite.bas
SECRET = "SEB$C0mpt4bl3#Pr0j3t9!"


def _hash_cle(texte: str) -> str:
    """Même algorithme que _HashCle() dans vba_Comptabilite.bas."""
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
    normalized = email.lower().strip()
    h = _hash_cle(normalized + SECRET)
    return f"SEB-{annee}-{h[:4]}-{h[4:]}"


def verifier_cle(cle: str, email: str) -> bool:
    """Vérifie qu'une clé est valide pour un email donné (test local)."""
    cle = cle.upper().strip()
    if len(cle) != 18:
        return False
    if not cle.startswith("SEB-"):
        return False
    hash_cle = cle[9:13] + cle[14:18]
    normalized = email.lower().strip()
    expected = _hash_cle(normalized + SECRET)
    return hash_cle == expected[:8]


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generer_cle.py email@exemple.com [annee]")
        sys.exit(1)

    email = sys.argv[1]
    annee = int(sys.argv[2]) if len(sys.argv) > 2 else None
    cle = generer_cle(email, annee)

    print()
    print(f"  Email : {email}")
    print(f"  Clé   : {cle}")
    print(f"  Vérif : {'OK' if verifier_cle(cle, email) else 'ERREUR'}")
    print()
    print("  -> Envoyez cette cle au donateur par email.")
    print("  -> Il la saisit dans EasyCompta apres son don PayPal.")
    print()
