# -*- coding: utf-8 -*-
"""Vérification du classeur : recalcul LibreOffice → zéro erreur de formule.

Convertit le .xlsx en Flat ODS (LibreOffice recalcule toutes les formules,
car openpyxl n'écrit aucune valeur en cache), puis inspecte chaque cellule
calculée à la recherche d'erreurs (#REF!, #DIV/0!, #VALUE!, #NAME?, …) et
affiche les indicateurs clés pour contrôle de cohérence.
"""

import re
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

SRC = Path("Gestion-SASU-SARL-2026.xlsx")
NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}
ERRORS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#N/A",
          "#NULL!", "Err:5")


def cell_text(cell):
    return "".join(cell.itertext())


def main():
    tmp = tempfile.mkdtemp()
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "fods", "--outdir", tmp,
         str(SRC)],
        check=True, capture_output=True, timeout=300)
    fods = Path(tmp) / (SRC.stem + ".fods")
    root = ET.parse(fods).getroot()

    problems = []
    n_formulas = 0
    values = {}
    for table in root.iter(f"{{{NS['table']}}}table"):
        sheet = table.get(f"{{{NS['table']}}}name")
        r = 0
        for row in table.iter(f"{{{NS['table']}}}table-row"):
            r += int(row.get(f"{{{NS['table']}}}number-rows-repeated", 1))
            c = 0
            for cell in row:
                if not cell.tag.endswith("}table-cell") and \
                   not cell.tag.endswith("}covered-table-cell"):
                    continue
                rep = int(cell.get(
                    f"{{{NS['table']}}}number-columns-repeated", 1))
                c += rep
                formula = cell.get(f"{{{NS['table']}}}formula")
                if formula:
                    n_formulas += 1
                txt = cell_text(cell)
                vtype = cell.get(f"{{{NS['office']}}}value-type")
                if vtype == "float":
                    values[(sheet, r, c)] = cell.get(
                        f"{{{NS['office']}}}value")
                for err in ERRORS:
                    if err in txt:
                        problems.append((sheet, r, c, formula, txt))

    print(f"Formules recalculées : {n_formulas}")
    if problems:
        print(f"\n❌ {len(problems)} erreur(s) de formule détectée(s) :")
        for p in problems[:40]:
            print("  ", p)
        sys.exit(1)
    print("✅ Aucune erreur de formule (#REF!, #DIV/0!, #VALUE!, #NAME?, "
          "#NUM!, #N/A).")

    # indicateurs clés (contrôle de cohérence) — feuille, ligne, colonne
    def show(labelled, sheet, r, c):
        v = values.get((sheet, r, c))
        print(f"  {labelled:<46} {v}")

    print("\nContrôles de cohérence :")
    show("Recettes — facturé HT annuel (O6)", "💰 Recettes", 6, 15)
    show("Dépenses — HT annuel (O6)", "💸 Dépenses", 6, 15)
    show("Suivi — CA réalisé cumul (D19)", "📊 Suivi mensuel", 19, 4)
    show("Suivi — dépenses réalisées cumul (G19)", "📊 Suivi mensuel", 19, 7)
    show("Budget — total recettes prévues (O12)", "📋 Budget", 12, 15)
    show("Budget — total dépenses prévues (O31)", "📋 Budget", 31, 15)
    show("Trésorerie — solde fin décembre (N20)", "🏦 Trésorerie", 20, 14)
    show("TVA — collectée annuelle (C19)", "🧾 TVA & Impôts", 19, 3)
    show("TVA — déductible annuelle (D19)", "🧾 TVA & Impôts", 19, 4)
    show("IS estimé (C30)", "🧾 TVA & Impôts", 30, 3)
    show("KM — total km (I7)", "🚗 Indemnités KM", 7, 9)
    show("KM — indemnité (I8)", "🚗 Indemnités KM", 8, 9)
    show("Dashboard — jauge budget % (K29)", "🏠 Dashboard", 29, 11)


if __name__ == "__main__":
    main()
