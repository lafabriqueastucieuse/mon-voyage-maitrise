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
                elif vtype == "string" and txt.strip():
                    values[(sheet, r, c)] = txt.strip()
                elif vtype == "date":
                    values[(sheet, r, c)] = cell.get(
                        f"{{{NS['office']}}}date-value")
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

    print("\nScore de santé & indicateurs (données de démo) :")
    for lab, r in [("Score trésorerie /100", 20),
                   ("Score rentabilité /100", 21),
                   ("Score budget /100", 22),
                   ("Score dynamique CA /100", 23),
                   ("Score provisions /100", 24),
                   ("SCORE GLOBAL /100", 25),
                   ("Runway (mois)", 11),
                   ("Marge nette", 12),
                   ("Seuil de rentabilité (€/mois)", 17),
                   ("Résultat projeté (run-rate)", 18)]:
        show(lab, "🏠 Dashboard", r, 17)
    show("Statut santé (D12)", "🏠 Dashboard", 12, 4)

    print("\nRecommandations (H13:H17) :")
    for r in range(13, 18):
        v = values.get(("🏠 Dashboard", r, 8))
        print(f"  {r}: {v}")

    print("\nÉchéancier :")
    show("Prochaine échéance (date)", "📅 Échéancier", 5, 10)
    show("Prochaine échéance (nature)", "📅 Échéancier", 6, 10)
    show("À prévoir ce mois-ci", "📅 Échéancier", 9, 10)
    show("Reste à payer sur l'année", "📅 Échéancier", 10, 10)

    print("\nSimulateurs :")
    show("Sim1 — CA HT à facturer/mois (C12)", "🧮 Simulateurs", 12, 3)
    show("Sim1 — TJM (C14)", "🧮 Simulateurs", 14, 3)
    show("Sim2 — net rémunération (C20)", "🧮 Simulateurs", 20, 3)
    show("Sim2 — net dividendes (C25)", "🧮 Simulateurs", 25, 3)
    show("Sim3 — résultat projeté hyp. (C37)", "🧮 Simulateurs", 37, 3)
    show("Sim3 — tréso fin d'année hyp. (C39)", "🧮 Simulateurs", 39, 3)


if __name__ == "__main__":
    main()
