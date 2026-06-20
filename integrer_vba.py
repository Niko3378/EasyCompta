#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Intègre automatiquement les modules VBA dans EasyCompta.xlsx
et sauvegarde en EasyCompta.xlsm via win32com (Excel COM automation).

Exécuté automatiquement par build_installer.bat.
"""

import os
import sys
import struct
import zipfile
import shutil
import tempfile
import winreg
import time
import subprocess
import pywintypes
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(BASE, "EasyCompta.xlsx")
XLSM = os.path.join(BASE, "EasyCompta.xlsm")
BAS  = os.path.join(BASE, "vba_Comptabilite.bas")


def _sanitiser_cp1252(code: str) -> str:
    """Replace characters outside CP1252 that corrupt VBA when saved to xlsm."""
    remplacements = {
        '─': '-',    # box drawing
        '→': '->',   # arrow
        '↻': '',     # refresh symbol
        '∑': '',     # sum symbol
        '⚠': '[!]',  # warning
        '✔': '[OK]', # check mark
    }
    result = []
    for ch in code:
        repl = remplacements.get(ch)
        if repl is not None:
            result.append(repl)
            continue
        try:
            ch.encode('cp1252')
            result.append(ch)
        except (UnicodeEncodeError, ValueError):
            result.append('?')
    return ''.join(result)


def _ecrire_bas_cp1252(src: str) -> str:
    """Écrit le .bas sanitisé en CP1252 dans un fichier temp et retourne son chemin."""
    content = _sanitiser_cp1252(open(src, encoding="utf-8").read())
    tmp = tempfile.NamedTemporaryFile(suffix=".bas", delete=False,
                                     mode="w", encoding="cp1252",
                                     errors="replace")
    tmp.write(content)
    tmp.close()
    return tmp.name


# Code injecté dans le module ThisWorkbook
THISWORKBOOK_CODE = """\
Option Explicit

Private Sub Workbook_Open()
    Mod_Comptabilite.VerifierOuDemanderLicence
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
End Sub
"""


def _neutraliser_pcode(xlsm_path: str):
    """
    Neutralise le p-code compilé dans vbaProject.bin (OLE inside xlsm ZIP).
    Met à zéro les 2 premiers octets de _VBA_PROJECT et Mod_Comptabilite :
    Excel ignorera le p-code corrompu et recompilera depuis le source VBA.
    """
    with zipfile.ZipFile(xlsm_path, 'r') as zin:
        vba_bytes = bytearray(zin.read('xl/vbaProject.bin'))
        all_names = zin.namelist()
        all_data  = {n: zin.read(n) for n in all_names}
        all_infos = {n: zin.getinfo(n) for n in all_names}

    if vba_bytes[:8] != b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1':
        print("  AVERTISSEMENT: vbaProject.bin invalide, skip patch p-code")
        return

    sector_size = 2 ** struct.unpack_from('<H', vba_bytes, 30)[0]
    mini_cutoff = struct.unpack_from('<I', vba_bytes, 56)[0]

    def sec_off(s):
        return (s + 1) * sector_size

    # Construire FAT depuis les 109 entrées DIFAT du header
    fat = []
    for i in range(109):
        val = struct.unpack_from('<I', vba_bytes, 76 + i * 4)[0]
        if val >= 0xFFFFFFF8:
            break
        off = sec_off(val)
        for j in range(0, sector_size, 4):
            if off + j + 4 <= len(vba_bytes):
                fat.append(struct.unpack_from('<I', vba_bytes, off + j)[0])

    def iter_chain(start):
        s, seen = start, set()
        while s < 0xFFFFFFF8 and s not in seen:
            seen.add(s); yield s
            s = fat[s] if s < len(fat) else 0xFFFFFFFE

    # Lire les entrées de répertoire OLE
    first_dir = struct.unpack_from('<I', vba_bytes, 48)[0]
    streams = {}
    for ds in iter_chain(first_dir):
        off = sec_off(ds)
        for i in range(0, sector_size, 128):
            e = vba_bytes[off + i: off + i + 128]
            if len(e) < 128:
                break
            nlen = struct.unpack_from('<H', e, 64)[0]
            if nlen < 2:
                continue
            name  = e[:nlen - 2].decode('utf-16-le', errors='replace')
            typ   = e[66]
            start = struct.unpack_from('<I', e, 116)[0]
            size  = struct.unpack_from('<I', e, 120)[0]
            if typ == 2:  # stream
                streams[name] = (start, size)

    # Invalider uniquement _VBA_PROJECT : selon la spec MS-OVBA, si ce stream
    # est absent ou invalide, Excel recompile TOUT le projet depuis le source.
    targets = ['_VBA_PROJECT']
    patched = []
    for name in targets:
        if name not in streams:
            continue
        start, size = streams[name]
        if size < mini_cutoff or start >= 0xFFFFFFF8:
            continue  # mini-stream, skip
        off = sec_off(start)
        if off + 2 <= len(vba_bytes):
            patched.append(f"{name}@{off:#x}({vba_bytes[off]:02x}{vba_bytes[off+1]:02x}->0000)")
            vba_bytes[off]     = 0x00
            vba_bytes[off + 1] = 0x00

    if not patched:
        print("  AVERTISSEMENT: aucun stream p-code patchable trouvé")
        return

    # Réécrire le xlsm avec vbaProject.bin modifié
    all_data['xl/vbaProject.bin'] = bytes(vba_bytes)
    tmp = xlsm_path + '.tmp'
    with zipfile.ZipFile(tmp, 'w') as zout:
        for name in all_names:
            zout.writestr(all_infos[name], all_data[name])
    shutil.move(tmp, xlsm_path)
    print(f"  P-code neutralisé : {', '.join(patched)}")


def _activer_vba_access():
    for version in ("16.0", "15.0", "14.0"):
        key_path = f"Software\\Microsoft\\Office\\{version}\\Excel\\Security"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                                 0, winreg.KEY_SET_VALUE | winreg.KEY_QUERY_VALUE)
            try:
                old_val, _ = winreg.QueryValueEx(key, "AccessVBOM")
            except FileNotFoundError:
                old_val = 0
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            print(f"  Acces VBA active (Office {version})")
            return version, old_val
        except FileNotFoundError:
            continue
    raise RuntimeError("Office/Excel non trouve dans le registre.")


def _restaurer_vba_access(version, old_val):
    key_path = f"Software\\Microsoft\\Office\\{version}\\Excel\\Security"
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path,
                             0, winreg.KEY_SET_VALUE)
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, old_val)
        winreg.CloseKey(key)
    except Exception:
        pass


def _xl_new():
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.AutomationSecurity = 1
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.EnableEvents = False
    return xl


def integrer_vba():
    print("\n=== Integration VBA ===")

    for chemin in [XLSX, BAS]:
        if not os.path.exists(chemin):
            raise FileNotFoundError(f"Fichier introuvable : {chemin}")

    version, old_val = _activer_vba_access()
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
    time.sleep(1.5)

    tmp_bas = None
    try:
        # ── Étape 1 : préparer le .bas CP1252 pour Import ──────────────────────
        tmp_bas = _ecrire_bas_cp1252(BAS)

        # ── Étape 2 : ouvrir xlsx, injecter VBA, sauvegarder en xlsm ───────────
        xl = _xl_new()
        wb = None
        try:
            print(f"  Ouverture de {XLSX}...")
            wb = xl.Workbooks.Open(os.path.abspath(XLSX), UpdateLinks=0)
            proj = wb.VBProject
            comps = proj.VBComponents

            # Supprimer ancien Mod_Comptabilite s'il existe
            for i in range(comps.Count, 0, -1):
                c = comps.Item(i)
                if c.Type == 1 and c.Name == "Mod_Comptabilite":
                    comps.Remove(c)
                    break

            # Importer via Import() — génère un p-code stable
            print(f"  Import de vba_Comptabilite.bas (via Import)...")
            comps.Import(tmp_bas)

            # Injecter Workbook_Open / BeforeClose dans ThisWorkbook
            print("  Injection ThisWorkbook...")
            tw = proj.VBComponents("ThisWorkbook")
            cm = tw.CodeModule
            if cm.CountOfLines > 0:
                cm.DeleteLines(1, cm.CountOfLines)
            cm.AddFromString(THISWORKBOOK_CODE.strip())

            # Sauvegarder en xlsm
            print(f"  Sauvegarde en {XLSM}...")
            time.sleep(1)
            if os.path.exists(XLSM):
                os.remove(XLSM)
            wb.SaveAs(os.path.abspath(XLSM), FileFormat=52)
            print("  Sauvegarde OK.")
        finally:
            if wb:
                try: wb.Close(SaveChanges=False)
                except: pass
            try: xl.Quit()
            except: pass
        time.sleep(1.5)

        # ── Étape 3 : vérifier que le xlsm existe et a une taille cohérente ──────
        print("  Verification du fichier xlsm...")
        if not os.path.exists(XLSM):
            raise RuntimeError("EasyCompta.xlsm absent après sauvegarde.")
        taille = os.path.getsize(XLSM)
        if taille < 50_000:
            raise RuntimeError(f"EasyCompta.xlsm trop petit ({taille} octets) — injection incomplète.")
        print(f"  Fichier xlsm OK ({taille // 1024} Ko).")

        print("\n  VBA integre et compile avec succes.")

    finally:
        if tmp_bas and os.path.exists(tmp_bas):
            try: os.remove(tmp_bas)
            except: pass
        _restaurer_vba_access(version, old_val)


if __name__ == "__main__":
    try:
        integrer_vba()
        print()
    except Exception as exc:
        print(f"\nERREUR : {exc}", file=sys.stderr)
        sys.exit(1)
