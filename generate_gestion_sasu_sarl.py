# -*- coding: utf-8 -*-
"""
Générateur du fichier « Gestion-SASU-SARL-2026.xlsx »
La Fabrique Astucieuse — outil premium de gestion financière SASU/SARL.

- 100 % formules Excel (aucune valeur calculée en dur), zéro macro.
- Compatible Excel + Google Sheets (fonctions standard, plages nommées).
- Palette : bordeaux #722F37 · terracotta #A94442 · taupe #8B7D6B ·
  beige #F5F0E8 · crème #FDFBF7 · vert sauge #7D8B6B.
"""

import datetime as dt

from openpyxl import Workbook
from openpyxl.chart import (BarChart, DoughnutChart, LineChart, Reference,
                            Series)
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.text import CharacterProperties
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

# ---------------------------------------------------------------- palette ---
BORDEAUX = "722F37"
TERRA = "A94442"
TAUPE = "8B7D6B"
TAUPE_L = "C9BFB0"      # bordures / filets
BEIGE = "F5F0E8"        # fond général
BEIGE_ALT = "EFE7DB"    # lignes alternées
CREME = "FDFBF7"        # cellules de saisie
SAUGE = "7D8B6B"
SAUGE_L = "E4E9DC"
INK = "3D3833"          # texte des cellules calculées
RED_L = "F6DBD9"        # alerte trésorerie < 0
ORANGE_L = "FBEAD5"     # alerte trésorerie < 1 000 €

FONT = "Montserrat"

# ---------------------------------------------------------------- formats ---
MONEY = '#,##0.00 "€";[Red](#,##0.00 "€")'
MONEY0 = '#,##0 "€";[Red](#,##0 "€")'
PCT = "0.0%"
PCT0 = "0%"
DATEF = "DD/MM/YYYY"
KMF = '#,##0 "km"'
NUMF = "#,##0.00"

MOIS = ["Janv", "Févr", "Mars", "Avr", "Mai", "Juin",
        "Juil", "Août", "Sept", "Oct", "Nov", "Déc"]

# noms d'onglets (ordre = ordre du classeur)
DASH = "🏠 Dashboard"
START = "🚀 Démarrer ici"
PAR = "⚙️ Paramètres"
BUD = "📋 Budget"
REC = "💰 Recettes"
DEP = "💸 Dépenses"
SUI = "📊 Suivi mensuel"
TRE = "🏦 Trésorerie"
TVA = "🧾 TVA & Impôts"
ECH = "📅 Échéancier"
SIM = "🧮 Simulateurs"
KM = "🚗 Indemnités KM"
LEX = "📖 Lexique"

MOIS_FR = ["janvier", "février", "mars", "avril", "mai", "juin",
           "juillet", "août", "septembre", "octobre", "novembre",
           "décembre"]

REC_FIRST, REC_LAST = 10, 209    # 200 lignes de saisie
DEP_FIRST, DEP_LAST = 10, 309    # 300 lignes
KM_FIRST, KM_LAST = 10, 159      # 150 lignes

# index du mois « en cours » borné à l'année de gestion
MIDX = "IF(YEAR(TODAY())>ANNEE,12,IF(YEAR(TODAY())<ANNEE,1,MONTH(TODAY())))"

# ----------------------------------------------------------------- styles ---


def fill(hexa):
    return PatternFill("solid", fgColor=hexa)


THIN = Side(style="thin", color=TAUPE_L)
THIN_T = Side(style="thin", color=TAUPE)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_INPUT = Border(left=THIN_T, right=THIN_T, top=THIN_T, bottom=THIN_T)

F_TITLE = Font(name=FONT, size=18, bold=True, color="FFFFFF")
F_SUB = Font(name=FONT, size=10, color="F5E6E0")
F_TIP = Font(name=FONT, size=9, italic=True, color=TAUPE)
F_BAND = Font(name=FONT, size=11, bold=True, color="FFFFFF")
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_SUBHEAD = Font(name=FONT, size=10, bold=True, color=TAUPE)
F_LABEL = Font(name=FONT, size=10, color=INK)
F_LABEL_B = Font(name=FONT, size=10, bold=True, color=INK)
F_INPUT = Font(name=FONT, size=10, color=BORDEAUX)
F_CALC = Font(name=FONT, size=10, color=INK)
F_CALC_B = Font(name=FONT, size=10, bold=True, color=INK)
F_TOTAL = Font(name=FONT, size=10, bold=True, color=BORDEAUX)
F_KPI_LAB = Font(name=FONT, size=8, bold=True, color=TAUPE)
F_KPI_VAL = Font(name=FONT, size=15, bold=True, color=BORDEAUX)
F_KPI_SUB = Font(name=FONT, size=7.5, italic=True, color=TAUPE)
F_NOTE = Font(name=FONT, size=8.5, italic=True, color=TAUPE)

A_C = Alignment(horizontal="center", vertical="center")
A_CW = Alignment(horizontal="center", vertical="center", wrap_text=True)
A_L = Alignment(horizontal="left", vertical="center")
A_LW = Alignment(horizontal="left", vertical="center", wrap_text=True)
A_R = Alignment(horizontal="right", vertical="center")

# ---------------------------------------------------------------- helpers ---


def paint(ws, max_col, max_row):
    """Fond beige + police par défaut sur toute la zone utile."""
    base = Font(name=FONT, size=10, color=INK)
    for row in ws.iter_rows(min_row=1, max_row=max_row, min_col=1,
                            max_col=max_col):
        for c in row:
            c.fill = fill(BEIGE)
            c.font = base
    ws.sheet_view.showGridLines = False


def merge(ws, rng, value=None, font=None, bg=None, align=None, border=None):
    first = rng.split(":")[0]
    for row in ws[rng]:
        for c in row:
            if font:
                c.font = font
            if bg:
                c.fill = fill(bg)
            if align:
                c.alignment = align
            if border:
                c.border = border
    ws.merge_cells(rng)
    if value is not None:
        ws[first] = value
    return ws[first]


def banner(ws, last_col, title, subtitle):
    merge(ws, f"A1:{last_col}1", title, F_TITLE, BORDEAUX,
          Alignment(horizontal="left", vertical="bottom", indent=1))
    merge(ws, f"A2:{last_col}2", subtitle, F_SUB, BORDEAUX,
          Alignment(horizontal="left", vertical="top", indent=1))
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 18
    # liseré terracotta sous le bandeau : la signature visuelle du fichier
    accent = Side(style="medium", color=TERRA)
    for c in ws[f"A2:{last_col}2"][0]:
        c.border = Border(bottom=accent)


def tip(ws, rng, text):
    merge(ws, rng, "💡 " + text, F_TIP, BEIGE, A_LW)
    ws.row_dimensions[int(rng.split(":")[0][1:])].height = 26


def band(ws, rng, text, bg=BORDEAUX):
    merge(ws, rng, text, F_BAND, bg,
          Alignment(horizontal="left", vertical="center", indent=1))
    ws.row_dimensions[int("".join(ch for ch in rng.split(":")[0]
                                  if ch.isdigit()))].height = 20


def head(ws, cell, text, bg=BORDEAUX):
    c = ws[cell]
    c.value = text
    c.font = F_HEAD
    c.fill = fill(bg)
    c.alignment = A_CW
    c.border = BORDER


def label(ws, cell, text, bold=False):
    c = ws[cell]
    c.value = text
    c.font = F_LABEL_B if bold else F_LABEL
    c.alignment = A_L


def cin(ws, cell, value=None, fmt=None):
    """Cellule de saisie : fond crème, texte bordeaux, bordure taupe."""
    c = ws[cell]
    if value is not None:
        c.value = value
    c.font = F_INPUT
    c.fill = fill(CREME)
    c.border = BORDER_INPUT
    c.alignment = A_R if fmt in (MONEY, MONEY0, PCT, PCT0, KMF, NUMF) else A_C
    if fmt:
        c.number_format = fmt
    return c


def ccalc(ws, cell, formula, fmt=MONEY, bold=False, bg=BEIGE_ALT, font=None):
    """Cellule calculée : fond beige, texte gris foncé."""
    c = ws[cell]
    c.value = formula
    c.font = font or (F_CALC_B if bold else F_CALC)
    c.fill = fill(bg)
    c.border = BORDER
    c.alignment = A_R
    c.number_format = fmt
    return c


def box(ws, r1, c1, r2, c2, color=TAUPE):
    """Contour fin autour d'un bloc (les bordures internes sont conservées)."""
    side = Side(style="thin", color=color)
    for cc in range(c1, c2 + 1):
        top = ws.cell(row=r1, column=cc)
        bot = ws.cell(row=r2, column=cc)
        top.border = Border(left=top.border.left, right=top.border.right,
                            top=side, bottom=top.border.bottom)
        bot.border = Border(left=bot.border.left, right=bot.border.right,
                            top=bot.border.top, bottom=side)
    for rr in range(r1, r2 + 1):
        lef = ws.cell(row=rr, column=c1)
        rig = ws.cell(row=rr, column=c2)
        lef.border = Border(left=side, right=lef.border.right,
                            top=lef.border.top, bottom=lef.border.bottom)
        rig.border = Border(left=rig.border.left, right=side,
                            top=rig.border.top, bottom=rig.border.bottom)


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def month_cols(first="C"):
    start = ord(first) - 64
    return [get_column_letter(start + i) for i in range(12)]


# ============================================================ DÉMARRER ICI ==

STEPS = [
    ("1", "Configurez votre entreprise",
     "Ouvrez l'onglet ⚙️ Paramètres et renseignez les 8 informations : nom, "
     "forme juridique, année, solde de départ, régime de TVA, statut du "
     "dirigeant, salaire net mensuel et CFE estimée. C'est le moteur de "
     "tout le fichier."),
    ("2", "Posez votre budget",
     "Remplissez le 📋 Budget prévisionnel, mois par mois. Astuce : pas "
     "d'idée précise ? Commencez large, vous ajusterez en cours d'année."),
    ("3", "Saisissez au fil de l'eau",
     "Chaque facture dans 💰 Recettes, chaque achat dans 💸 Dépenses. "
     "Règle d'or : 5 minutes par semaine suffisent pour rester à jour."),
    ("4", "Suivez vos trajets",
     "À chaque déplacement professionnel, notez la date, le motif et les "
     "kilomètres dans 🚗 Indemnités KM. Ce sont des remboursements que "
     "beaucoup oublient."),
    ("5", "Pilotez",
     "Ouvrez le 🏠 Dashboard : score de santé, alertes, recommandations et "
     "📅 Échéancier font le reste. Vous savez toujours où vous en êtes."),
]

ROUTINE = [
    "◻  Toutes les factures et toutes les dépenses du mois sont saisies.",
    "◻  Les encaissements sont pointés (colonne « Encaissé » des 💰 "
    "Recettes).",
    "◻  La TVA du mois est vérifiée dans 🧾 TVA & Impôts.",
    "◻  Les provisions (URSSAF, IS, CFE) sont disponibles sur le compte.",
    "◻  Les écarts du 📊 Suivi mensuel sont analysés (favorable ou non ?).",
    "◻  La trésorerie du mois prochain est vérifiée dans 🏦 Trésorerie.",
]


def build_start(ws):
    paint(ws, 11, 36)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 6, "I": 13, "J": 3})
    for cl in "CDEFGH":
        ws.column_dimensions[cl].width = 13
    banner(ws, "J", "🚀 DÉMARRER ICI",
           "Bienvenue dans votre assistant de pilotage. En 15 minutes, "
           "votre entreprise est configurée.")
    tip(ws, "B3:J3",
        "Suivez les 5 étapes dans l'ordre, cochez-les au fur et à mesure "
        "(menu déroulant à droite), puis revenez ici chaque fin de mois "
        "pour votre routine de 10 minutes.")

    # barre de progression de la configuration
    prog = merge(ws, "B4:I4",
                 '="Configuration : "&COUNTIF($I$7:$I$19,"✅")'
                 '&"/5 étapes ✅"',
                 Font(name=FONT, size=13, bold=True, color=BORDEAUX),
                 CREME, A_C, BORDER_INPUT)
    ws.row_dimensions[4].height = 26

    dv_check = DataValidation(type="list", formula1='"☐,✅"',
                              allow_blank=True)
    ws.add_data_validation(dv_check)

    r = 7
    for num, titre, desc in STEPS:
        c = ws[f"B{r}"]
        c.value = num
        c.font = Font(name=FONT, size=16, bold=True, color=BORDEAUX)
        c.fill = fill(CREME)
        c.alignment = A_C
        merge(ws, f"C{r}:H{r}", titre,
              Font(name=FONT, size=12, bold=True, color=BORDEAUX), CREME,
              A_L)
        chk = cin(ws, f"I{r}", "☐")
        chk.font = Font(name=FONT, size=12, color=BORDEAUX)
        dv_check.add(f"I{r}")
        merge(ws, f"B{r + 1}:B{r + 1}", "", F_LABEL, CREME)
        merge(ws, f"C{r + 1}:I{r + 1}", desc, F_LABEL, CREME, A_LW)
        ws.row_dimensions[r].height = 22
        ws.row_dimensions[r + 1].height = 30
        box(ws, r, 2, r + 1, 9, TAUPE_L)
        r += 3

    # routine mensuelle
    band(ws, "B22:I22", "⏱️ VOTRE ROUTINE MENSUELLE EN 10 MINUTES")
    for i, item in enumerate(ROUTINE):
        merge(ws, f"B{23 + i}:I{23 + i}", item, F_LABEL, BEIGE_ALT, A_LW,
              BORDER)
        ws.row_dimensions[23 + i].height = 18
    box(ws, 22, 2, 28, 9, TAUPE)

    # renvoi vers le lexique
    merge(ws, "B30:I31",
          "❓ Une question ? L'onglet 📖 Lexique vous explique tous les "
          "termes (TVA, IS, runway, provision…) en langage simple, avec "
          "des exemples concrets.",
          Font(name=FONT, size=10, italic=True, color=TAUPE), CREME, A_LW)
    box(ws, 30, 2, 31, 9, TAUPE_L)


# ============================================================== PARAMÈTRES ==

def build_parametres(ws):
    paint(ws, 12, 40)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 34, "C": 22, "D": 3, "E": 26, "F": 17,
                "G": 19, "H": 21, "I": 15, "J": 3})
    banner(ws, "J", "⚙️ PARAMÈTRES",
           "En clair : la fiche d'identité de votre gestion — remplie une "
           "fois, utilisée par tous les onglets.")
    tip(ws, "B3:J3",
        "Remplissez uniquement les cellules crème. Elles alimentent "
        "automatiquement tous les autres onglets : ne modifiez ni l'ordre "
        "des listes ni le barème sans raison.")

    # --- identité ------------------------------------------------------
    band(ws, "B5:C5", "Identité de l'entreprise")
    label(ws, "B6", "Nom de l'entreprise")
    cin(ws, "C6", "Astuce Conseil")
    label(ws, "B7", "Forme juridique")
    cin(ws, "C7", "SASU")
    label(ws, "B8", "Année de gestion")
    cin(ws, "C8", 2026, "0")
    label(ws, "B9", "Solde de trésorerie au 01/01")
    cin(ws, "C9", 12500, MONEY)

    # --- fiscalité -----------------------------------------------------
    band(ws, "B11:C11", "TVA & fiscalité")
    label(ws, "B12", "Régime de TVA")
    cin(ws, "C12", "Réel normal")
    label(ws, "B13", "Taux de TVA par défaut")
    cin(ws, "C13", 0.20, PCT0)
    label(ws, "B14", "CFE estimée (annuelle)")
    cin(ws, "C14", 650, MONEY)

    # --- dirigeant -----------------------------------------------------
    band(ws, "B16:C16", "Dirigeant & cotisations sociales")
    label(ws, "B17", "Statut du dirigeant")
    cin(ws, "C17", "Assimilé salarié (SASU)")
    label(ws, "B18", "Taux de charges — Assimilé salarié (SASU)")
    cin(ws, "C18", 0.82, PCT0)
    label(ws, "B19", "Taux de charges — TNS gérant majoritaire (SARL)")
    cin(ws, "C19", 0.45, PCT0)
    label(ws, "B20", "Salaire net mensuel versé")
    cin(ws, "C20", 2600, MONEY)
    label(ws, "B21", "Taux de charges appliqué (auto, selon le statut)")
    ccalc(ws, "C21",
          '=IF($C$17="Assimilé salarié (SASU)",$C$18,$C$19)', PCT0)
    merge(ws, "B22:C22",
          "Taux moyens indicatifs, à ajuster selon votre situation.",
          F_NOTE, BEIGE, A_LW)

    # --- IS ------------------------------------------------------------
    band(ws, "E5:F5", "Impôt sur les sociétés (taux modifiables)")
    label(ws, "E6", "Seuil du taux réduit")
    cin(ws, "F6", 42500, MONEY0)
    label(ws, "E7", "Taux réduit (PME)")
    cin(ws, "F7", 0.15, PCT0)
    label(ws, "E8", "Taux normal")
    cin(ws, "F8", 0.25, PCT0)
    merge(ws, "E9:F9",
          "Taux réduit sous conditions PME (CA < 10 M€, capital détenu à "
          "75 % par des personnes physiques) — valeurs indicatives.",
          F_NOTE, BEIGE, A_LW)
    ws.row_dimensions[9].height = 34

    # --- barème kilométrique -------------------------------------------
    band(ws, "E11:I11", "Barème kilométrique — voitures (modifiable)")
    for col, txt in zip("EFGHI", ["Puissance fiscale", "≤ 5 000 km (×)",
                                  "5 001 à 20 000 km (×)",
                                  "5 001 à 20 000 km (+)",
                                  "> 20 000 km (×)"]):
        head(ws, f"{col}12", txt, TAUPE)
    ws.row_dimensions[12].height = 30
    bareme = [("3 CV et moins", 0.529, 0.316, 1065, 0.370),
              ("4 CV", 0.606, 0.340, 1330, 0.407),
              ("5 CV", 0.636, 0.357, 1395, 0.427),
              ("6 CV", 0.665, 0.374, 1457, 0.447),
              ("7 CV et plus", 0.697, 0.394, 1515, 0.470)]
    for i, row in enumerate(bareme):
        r = 13 + i
        c = ws[f"E{r}"]
        c.value = row[0]
        c.font = F_LABEL_B
        c.fill = fill(BEIGE_ALT)
        c.border = BORDER
        c.alignment = A_L
        cin(ws, f"F{r}", row[1], "0.000")
        cin(ws, f"G{r}", row[2], "0.000")
        cin(ws, f"H{r}", row[3], MONEY0)
        cin(ws, f"I{r}", row[4], "0.000")

    # --- échéancier & simulateurs ---------------------------------------
    band(ws, "B24:C24", "Échéancier & simulateurs")
    label(ws, "B25", "Jour de prélèvement URSSAF (5 ou 15)")
    cin(ws, "C25", 5, "0")
    label(ws, "B26", "Flat tax sur dividendes (PFU)")
    cin(ws, "C26", 0.30, PCT0)
    label(ws, "B27", "Acompte TVA de juillet (réel simplifié)")
    cin(ws, "C27", 0.55, PCT0)
    label(ws, "B28", "Acompte TVA de décembre (réel simplifié)")
    cin(ws, "C28", 0.40, PCT0)
    merge(ws, "B29:C29",
          "Jours et taux usuels, à ajuster selon votre situation.",
          F_NOTE, BEIGE, A_LW)
    dv_jour = DataValidation(type="list", formula1='"5,15"',
                             allow_blank=True)
    ws.add_data_validation(dv_jour)
    dv_jour.add("C25")

    # --- listes de référence -------------------------------------------
    band(ws, "E20:I20", "Listes de référence (menus déroulants — ne pas "
                        "déplacer)", TAUPE)
    label(ws, "E21", "Taux de TVA")
    for i, rate in enumerate([0.20, 0.10, 0.055, 0.021, 0]):
        cin(ws, f"F{21 + i}", rate, PCT)
    merge(ws, "G21:I22",
          "Les catégories de dépenses des menus déroulants sont celles de "
          "l'onglet 📋 Budget (colonne Catégorie).", F_NOTE, BEIGE, A_LW)

    # menus déroulants
    dv_forme = DataValidation(type="list", formula1='"SASU,SARL"',
                              allow_blank=True)
    dv_regime = DataValidation(
        type="list",
        formula1='"Franchise en base,Réel simplifié,Réel normal"',
        allow_blank=True)
    dv_statut = DataValidation(
        type="list",
        formula1='"Assimilé salarié (SASU),TNS gérant majoritaire (SARL)"',
        allow_blank=True)
    for dv, cell in [(dv_forme, "C7"), (dv_regime, "C12"),
                     (dv_statut, "C17")]:
        ws.add_data_validation(dv)
        dv.add(cell)


def define_names(wb):
    q = f"'{PAR}'!"
    names = {
        "NOM_ENT": f"{q}$C$6",
        "ANNEE": f"{q}$C$8",
        "SOLDE_INIT": f"{q}$C$9",
        "REGIME_TVA": f"{q}$C$12",
        "TVA_DEFAUT": f"{q}$C$13",
        "CFE_ANNUELLE": f"{q}$C$14",
        "STATUT": f"{q}$C$17",
        "TAUX_CHG_AS": f"{q}$C$18",
        "TAUX_CHG_TNS": f"{q}$C$19",
        "SALAIRE_NET": f"{q}$C$20",
        "TAUX_CHARGES": f"{q}$C$21",
        "SEUIL_IS": f"{q}$F$6",
        "TAUX_IS_REDUIT": f"{q}$F$7",
        "TAUX_IS_NORMAL": f"{q}$F$8",
        "JOUR_URSSAF": f"{q}$C$25",
        "FLAT_TAX": f"{q}$C$26",
        "TVA_AC_JUIL": f"{q}$C$27",
        "TVA_AC_DEC": f"{q}$C$28",
        "BAREME_KM": f"{q}$E$13:$I$17",
        "LISTE_TVA": f"{q}$F$21:$F$25",
        "LISTE_CATEGORIES": f"'{BUD}'!$B$16:$B$30",
    }
    for name, ref in names.items():
        wb.defined_names[name] = DefinedName(name, attr_text=ref)


# ================================================================== BUDGET ==

REC_CATS = ["Prestations de conseil", "Formations", "Ateliers & séminaires",
            "Commissions & apport d'affaires", "Autres recettes"]
DEP_CATS = ["Achats & fournitures", "Sous-traitance", "Loyer & charges",
            "Assurances", "Télécom & internet", "Logiciels & abonnements",
            "Marketing & communication", "Déplacements & missions",
            "Salaire net dirigeant", "Charges sociales dirigeant",
            "Honoraires comptables", "Frais bancaires", "Formation",
            "Entretien & petit équipement", "Divers & imprévus"]

BUD_REC = {  # valeurs prévisionnelles d'exemple
    0: [6500, 7000, 7000, 7500, 7500, 8000, 6000, 4500, 8000, 8500, 8500, 7500],
    1: [1200, 0, 2400, 1200, 0, 2400, 0, 0, 2400, 1200, 0, 1200],
    2: [0, 800, 0, 800, 0, 800, 0, 0, 800, 0, 800, 0],
    3: [150] * 12,
    4: [0] * 12,
}
BUD_DEP = {
    0: [120] * 12, 1: [400, 400, 600, 400, 400, 600, 200, 0, 600, 400, 400, 600],
    2: [350] * 12, 3: [45] * 12, 4: [55] * 12, 5: [130] * 12,
    6: [250] * 12, 7: [180] * 12,
    10: [120] * 12, 11: [29] * 12,
    12: [0, 0, 0, 800, 0, 0, 0, 0, 0, 600, 0, 0],
    13: [40] * 12, 14: [100] * 12,
}


def build_budget(ws):
    paint(ws, 16, 40)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 30, "O": 13, "P": 3})
    for col in month_cols():
        ws.column_dimensions[col].width = 10.5
    banner(ws, "O", "📋 BUDGET PRÉVISIONNEL",
           "En clair : ce que vous prévoyez de gagner et de dépenser — "
           "votre feuille de route de l'année.")
    tip(ws, "B3:O3",
        "Saisissez vos prévisions dans les cellules crème (catégories "
        "personnalisables). Le salaire du dirigeant et ses charges sont "
        "pré-remplis depuis l'onglet ⚙️ Paramètres — vous pouvez les écraser.")

    cols = month_cols()

    def header_row(r):
        head(ws, f"B{r}", "Catégorie", TAUPE)
        for i, cl in enumerate(cols):
            head(ws, f"{cl}{r}", MOIS[i], TAUPE)
        head(ws, f"O{r}", "Total", TAUPE)

    # --- recettes -------------------------------------------------------
    band(ws, "B5:O5", "RECETTES PRÉVUES")
    header_row(6)
    for i, cat in enumerate(REC_CATS):
        r = 7 + i
        cin(ws, f"B{r}", cat).alignment = A_L
        for m, cl in enumerate(cols):
            cin(ws, f"{cl}{r}", BUD_REC.get(i, [0] * 12)[m], MONEY0)
        ccalc(ws, f"O{r}", f"=SUM(C{r}:N{r})", MONEY0)
    label(ws, "B12", "Total recettes prévues", bold=True)
    ws["B12"].font = F_TOTAL
    for cl in cols + ["O"]:
        ccalc(ws, f"{cl}12", f"=SUM({cl}7:{cl}11)", MONEY0, bold=True,
              font=F_TOTAL)

    # --- dépenses ---------------------------------------------------------
    band(ws, "B14:O14", "DÉPENSES PRÉVUES")
    header_row(15)
    for i, cat in enumerate(DEP_CATS):
        r = 16 + i
        cin(ws, f"B{r}", cat).alignment = A_L
        for m, cl in enumerate(cols):
            if i == 8:      # salaire net dirigeant
                cin(ws, f"{cl}{r}", "=SALAIRE_NET", MONEY0)
            elif i == 9:    # charges sociales dirigeant
                cin(ws, f"{cl}{r}", "=ROUND(SALAIRE_NET*TAUX_CHARGES,0)",
                    MONEY0)
            else:
                cin(ws, f"{cl}{r}", BUD_DEP.get(i, [0] * 12)[m], MONEY0)
        ccalc(ws, f"O{r}", f"=SUM(C{r}:N{r})", MONEY0)
    label(ws, "B31", "Total dépenses prévues", bold=True)
    ws["B31"].font = F_TOTAL
    for cl in cols + ["O"]:
        ccalc(ws, f"{cl}31", f"=SUM({cl}16:{cl}30)", MONEY0, bold=True,
              font=F_TOTAL)

    # --- résultat ---------------------------------------------------------
    label(ws, "B33", "Résultat prévisionnel mensuel", bold=True)
    label(ws, "B34", "Résultat prévisionnel cumulé", bold=True)
    for i, cl in enumerate(cols):
        ccalc(ws, f"{cl}33", f"={cl}12-{cl}31", MONEY0, bold=True)
        prev = f"{cols[i - 1]}34+" if i else ""
        ccalc(ws, f"{cl}34", f"={prev}{cl}33", MONEY0, bold=True)
    ccalc(ws, "O33", "=O12-O31", MONEY0, bold=True)
    ccalc(ws, "O34", "=N34", MONEY0, bold=True)
    ws.freeze_panes = "C7"


# ======================================================= RECETTES/DÉPENSES ==

REC_SAMPLES = [
    (dt.date(2026, 1, 8), "Studio Novelli", "Audit organisation & process",
     3200, 0.20, "Oui"),
    (dt.date(2026, 1, 15), "Boulangerie Martin",
     "Accompagnement digitalisation", 1450, 0.20, "Oui"),
    (dt.date(2026, 1, 22), "Mairie de Beaulieu",
     "Formation gestion de projet (2 j)", 2400, 0.20, "Oui"),
    (dt.date(2026, 1, 30), "Cabinet Ferrand",
     "Conseil stratégie — forfait janvier", 1800, 0.20, "Oui"),
    (dt.date(2026, 2, 6), "Studio Novelli", "Suivi mensuel — février",
     950, 0.20, "Oui"),
    (dt.date(2026, 2, 12), "Éditions Clairval",
     "Atelier productivité équipe", 1600, 0.20, "Oui"),
    (dt.date(2026, 2, 20), "Boutique Léonie", "Refonte parcours client",
     2750, 0.20, "Non"),
    (dt.date(2026, 2, 27), "Cabinet Ferrand",
     "Conseil stratégie — forfait février", 1800, 0.20, "Oui"),
    (dt.date(2026, 3, 10), "Groupe Vaillant",
     "Diagnostic flash + restitution", 3900, 0.20, "Non"),
    (dt.date(2026, 3, 18), "Mairie de Beaulieu",
     "Formation bureautique (1 j)", 1200, 0.20, "Non"),
]

DEP_SAMPLES = [
    (dt.date(2026, 1, 5), "OVHcloud", "Logiciels & abonnements",
     "Hébergement site + e-mails", 14.90, 0.20, "Oui"),
    (dt.date(2026, 1, 5), "AXA Pro", "Assurances",
     "RC Pro — prime mensuelle", 38.50, 0.0, "Oui"),
    (dt.date(2026, 1, 9), "SNCF Connect", "Déplacements & missions",
     "AR Paris — réunion client", 89.00, 0.10, "Oui"),
    (dt.date(2026, 1, 12), "Adobe", "Logiciels & abonnements",
     "Creative Cloud", 59.99, 0.20, "Oui"),
    (dt.date(2026, 1, 15), "Compta Facile", "Honoraires comptables",
     "Forfait comptable janvier", 120.00, 0.20, "Oui"),
    (dt.date(2026, 1, 20), "Orange Pro", "Télécom & internet",
     "Forfait mobile + fibre", 54.90, 0.20, "Oui"),
    (dt.date(2026, 1, 28), "Qonto", "Frais bancaires",
     "Abonnement compte pro", 29.00, 0.20, "Oui"),
    (dt.date(2026, 2, 3), "Bureau Vallée", "Achats & fournitures",
     "Papeterie & consommables", 47.60, 0.20, "Oui"),
    (dt.date(2026, 2, 10), "Meta Ads", "Marketing & communication",
     "Campagne prospection février", 150.00, 0.20, "Oui"),
    (dt.date(2026, 2, 18), "La Poste", "Achats & fournitures",
     "Affranchissements", 21.40, 0.20, "Non"),
]


def synth_block(ws, label_ht, label_ttc, f_ht, f_ttc):
    """Bloc « totaux par mois » au-dessus des tableaux de saisie."""
    cols = month_cols()
    head(ws, "B5", "Synthèse mensuelle", TAUPE)
    for i, cl in enumerate(cols):
        head(ws, f"{cl}5", MOIS[i], TAUPE)
    head(ws, "O5", "Année", TAUPE)
    label(ws, "B6", label_ht, bold=True)
    label(ws, "B7", label_ttc, bold=True)
    for m, cl in enumerate(cols):
        ccalc(ws, f"{cl}6", f_ht.format(m=m + 1), MONEY0)
        ccalc(ws, f"{cl}7", f_ttc.format(m=m + 1), MONEY0)
    ccalc(ws, "O6", "=SUM(C6:N6)", MONEY0, bold=True)
    ccalc(ws, "O7", "=SUM(C7:N7)", MONEY0, bold=True)


def build_recettes(ws):
    paint(ws, 16, REC_LAST + 6)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 12, "C": 22, "D": 34, "E": 13, "F": 9.5,
                "G": 12, "H": 13, "I": 11, "J": 7.5, "K": 3,
                "L": 10.5, "M": 10.5, "N": 10.5, "O": 12})
    banner(ws, "O", "💰 RECETTES",
           "En clair : tout ce que vous facturez à vos clients — la source "
           "de votre chiffre d'affaires.")
    tip(ws, "B3:O3",
        "Une ligne par facture : remplissez les cellules crème (date, "
        "client, description, montant HT, taux de TVA, encaissé). Les dix "
        "premières lignes sont des exemples à remplacer par vos données.")

    synth_block(
        ws, "Facturé HT", "Encaissé TTC",
        "=SUMIFS($E${f}:$E${l},$J${f}:$J${l},{{m}})".format(
            f=REC_FIRST, l=REC_LAST),
        "=SUMIFS($H${f}:$H${l},$J${f}:$J${l},{{m}},$I${f}:$I${l},\"Oui\")"
        .format(f=REC_FIRST, l=REC_LAST))

    headers = ["Date", "Client", "Description", "Montant HT", "Taux TVA",
               "TVA", "Montant TTC", "Encaissé", "Mois"]
    for i, h in enumerate(headers):
        head(ws, f"{get_column_letter(2 + i)}9", h)
    ws.row_dimensions[9].height = 22

    for r in range(REC_FIRST, REC_LAST + 1):
        alt = BEIGE_ALT if (r - REC_FIRST) % 2 else BEIGE
        cin(ws, f"B{r}", fmt=DATEF)
        cin(ws, f"C{r}").alignment = A_L
        cin(ws, f"D{r}").alignment = A_L
        cin(ws, f"E{r}", fmt=MONEY)
        cin(ws, f"F{r}", fmt=PCT)
        ccalc(ws, f"G{r}",
              f'=IF(OR($E{r}="",$F{r}=""),"",ROUND($E{r}*$F{r},2))',
              MONEY, bg=alt)
        ccalc(ws, f"H{r}",
              f'=IF($E{r}="","",ROUND($E{r}*(1+IF($F{r}="",0,$F{r})),2))',
              MONEY, bg=alt)
        cin(ws, f"I{r}")
        ccalc(ws, f"J{r}", f'=IF($B{r}="","",MONTH($B{r}))', "0", bg=alt)

    for i, s in enumerate(REC_SAMPLES):
        r = REC_FIRST + i
        ws[f"B{r}"] = s[0]
        ws[f"C{r}"] = s[1]
        ws[f"D{r}"] = s[2]
        ws[f"E{r}"] = s[3]
        ws[f"F{r}"] = s[4]
        ws[f"I{r}"] = s[5]

    dv_tva = DataValidation(type="list", formula1="LISTE_TVA",
                            allow_blank=True)
    dv_enc = DataValidation(type="list", formula1='"Oui,Non"',
                            allow_blank=True)
    ws.add_data_validation(dv_tva)
    ws.add_data_validation(dv_enc)
    dv_tva.add(f"F{REC_FIRST}:F{REC_LAST}")
    dv_enc.add(f"I{REC_FIRST}:I{REC_LAST}")
    ws.freeze_panes = "A10"


def build_depenses(ws):
    paint(ws, 16, DEP_LAST + 6)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 12, "C": 20, "D": 26, "E": 30, "F": 13,
                "G": 9.5, "H": 13, "I": 13, "J": 10, "K": 7.5, "L": 3,
                "M": 10.5, "N": 10.5, "O": 12})
    banner(ws, "O", "💸 DÉPENSES",
           "En clair : tout ce que l'entreprise dépense — pour garder la "
           "main sur vos coûts.")
    tip(ws, "B3:O3",
        "Choisissez la catégorie dans le menu déroulant (liste de l'onglet "
        "📋 Budget). Ne saisissez pas ici le salaire du dirigeant, les "
        "cotisations, la TVA ou les impôts : ils sont gérés automatiquement "
        "dans 🏦 Trésorerie et 🧾 TVA & Impôts.")

    # bloc synthèse (colonnes C..N du bloc = mois)
    cols = month_cols()
    head(ws, "B5", "Synthèse mensuelle", TAUPE)
    for i, cl in enumerate(cols):
        head(ws, f"{cl}5", MOIS[i], TAUPE)
    head(ws, "O5", "Année", TAUPE)
    label(ws, "B6", "Dépenses HT", bold=True)
    label(ws, "B7", "Payées TTC", bold=True)
    for m, cl in enumerate(cols):
        ccalc(ws, f"{cl}6",
              f"=SUMIFS($F${DEP_FIRST}:$F${DEP_LAST},"
              f"$K${DEP_FIRST}:$K${DEP_LAST},{m + 1})", MONEY0)
        ccalc(ws, f"{cl}7",
              f"=SUMIFS($I${DEP_FIRST}:$I${DEP_LAST},"
              f"$K${DEP_FIRST}:$K${DEP_LAST},{m + 1},"
              f"$J${DEP_FIRST}:$J${DEP_LAST},\"Oui\")", MONEY0)
    ccalc(ws, "O6", "=SUM(C6:N6)", MONEY0, bold=True)
    ccalc(ws, "O7", "=SUM(C7:N7)", MONEY0, bold=True)

    headers = ["Date", "Fournisseur", "Catégorie", "Description",
               "Montant HT", "Taux TVA", "TVA déductible", "Montant TTC",
               "Payé", "Mois"]
    for i, h in enumerate(headers):
        head(ws, f"{get_column_letter(2 + i)}9", h)
    ws.row_dimensions[9].height = 22

    for r in range(DEP_FIRST, DEP_LAST + 1):
        alt = BEIGE_ALT if (r - DEP_FIRST) % 2 else BEIGE
        cin(ws, f"B{r}", fmt=DATEF)
        cin(ws, f"C{r}").alignment = A_L
        cin(ws, f"D{r}").alignment = A_L
        cin(ws, f"E{r}").alignment = A_L
        cin(ws, f"F{r}", fmt=MONEY)
        cin(ws, f"G{r}", fmt=PCT)
        ccalc(ws, f"H{r}",
              f'=IF(OR($F{r}="",$G{r}=""),"",ROUND($F{r}*$G{r},2))',
              MONEY, bg=alt)
        ccalc(ws, f"I{r}",
              f'=IF($F{r}="","",ROUND($F{r}*(1+IF($G{r}="",0,$G{r})),2))',
              MONEY, bg=alt)
        cin(ws, f"J{r}")
        ccalc(ws, f"K{r}", f'=IF($B{r}="","",MONTH($B{r}))', "0", bg=alt)

    for i, s in enumerate(DEP_SAMPLES):
        r = DEP_FIRST + i
        ws[f"B{r}"] = s[0]
        ws[f"C{r}"] = s[1]
        ws[f"D{r}"] = s[2]
        ws[f"E{r}"] = s[3]
        ws[f"F{r}"] = s[4]
        ws[f"G{r}"] = s[5]
        ws[f"J{r}"] = s[6]

    dv_cat = DataValidation(type="list", formula1="LISTE_CATEGORIES",
                            allow_blank=True)
    dv_tva = DataValidation(type="list", formula1="LISTE_TVA",
                            allow_blank=True)
    dv_pay = DataValidation(type="list", formula1='"Oui,Non"',
                            allow_blank=True)
    for dv, rng in [(dv_cat, f"D{DEP_FIRST}:D{DEP_LAST}"),
                    (dv_tva, f"G{DEP_FIRST}:G{DEP_LAST}"),
                    (dv_pay, f"J{DEP_FIRST}:J{DEP_LAST}")]:
        ws.add_data_validation(dv)
        dv.add(rng)
    ws.freeze_panes = "A10"


# =========================================================== SUIVI MENSUEL ==

def build_suivi(ws):
    paint(ws, 13, 26)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 11})
    for cl in "CDEFGHIJK":
        ws.column_dimensions[cl].width = 13.5
    widths(ws, {"L": 3})
    banner(ws, "K", "📊 SUIVI MENSUEL — PRÉVISIONNEL vs RÉALISÉ",
           "En clair : prévu contre réalisé — pour savoir si vous êtes sur "
           "la bonne trajectoire.")
    tip(ws, "B3:K3",
        "Un écart positif est toujours favorable (vert) ; un écart négatif "
        "est défavorable (rouge). Les dépenses réalisées incluent "
        "automatiquement le salaire net + cotisations du dirigeant pour les "
        "mois écoulés.")

    band(ws, "C5:E5", "RECETTES")
    band(ws, "F5:H5", "DÉPENSES", TAUPE)
    band(ws, "I5:K5", "RÉSULTAT")
    head(ws, "B6", "Mois", TAUPE)
    for grp_start in ("C", "F", "I"):
        base = ord(grp_start)
        head(ws, f"{chr(base)}6", "Prévu", TAUPE)
        head(ws, f"{chr(base + 1)}6", "Réalisé", TAUPE)
        head(ws, f"{chr(base + 2)}6", "Écart", TAUPE)

    for m in range(1, 13):
        r = 6 + m
        bl = get_column_letter(2 + m)   # colonne du mois dans 📋 Budget
        c = ws[f"B{r}"]
        c.value = MOIS[m - 1]
        c.font = F_LABEL_B
        c.fill = fill(BEIGE_ALT)
        c.border = BORDER
        c.alignment = A_C
        alt = BEIGE_ALT if m % 2 else BEIGE
        ccalc(ws, f"C{r}", f"='{BUD}'!{bl}12", MONEY0, bg=alt)
        ccalc(ws, f"D{r}",
              f"=SUMIFS('{REC}'!$E${REC_FIRST}:$E${REC_LAST},"
              f"'{REC}'!$J${REC_FIRST}:$J${REC_LAST},{m})", MONEY0, bg=alt)
        ccalc(ws, f"E{r}", f"=D{r}-C{r}", MONEY0, bg=alt)
        ccalc(ws, f"F{r}", f"='{BUD}'!{bl}31", MONEY0, bg=alt)
        ccalc(ws, f"G{r}",
              f"=SUMIFS('{DEP}'!$F${DEP_FIRST}:$F${DEP_LAST},"
              f"'{DEP}'!$K${DEP_FIRST}:$K${DEP_LAST},{m})"
              f"+IF(OR(YEAR(TODAY())>ANNEE,AND(YEAR(TODAY())=ANNEE,"
              f"MONTH(TODAY())>={m})),ROUND(SALAIRE_NET*(1+TAUX_CHARGES),2)"
              f",0)", MONEY0, bg=alt)
        ccalc(ws, f"H{r}", f"=F{r}-G{r}", MONEY0, bg=alt)
        ccalc(ws, f"I{r}", f"=C{r}-F{r}", MONEY0, bg=alt)
        ccalc(ws, f"J{r}", f"=D{r}-G{r}", MONEY0, bg=alt)
        ccalc(ws, f"K{r}", f"=J{r}-I{r}", MONEY0, bg=alt)

    r = 19
    c = ws[f"B{r}"]
    c.value = "CUMUL"
    c.font = F_HEAD
    c.fill = fill(BORDEAUX)
    c.border = BORDER
    c.alignment = A_C
    for cl in "CDEFGHIJK":
        ccalc(ws, f"{cl}{r}", f"=SUM({cl}7:{cl}18)", MONEY0, bold=True,
              font=F_TOTAL)

    # écarts : vert sauge si favorable, terracotta sinon
    sauge_f = Font(name=FONT, size=10, bold=True, color=SAUGE)
    terra_f = Font(name=FONT, size=10, bold=True, color=TERRA)
    for rng in ("E7:E19", "H7:H19", "K7:K19"):
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="greaterThanOrEqual", formula=["0"], font=sauge_f))
        ws.conditional_formatting.add(rng, CellIsRule(
            operator="lessThan", formula=["0"], font=terra_f))
    ws.freeze_panes = "C7"


# ============================================================== TRÉSORERIE ==

def build_tresorerie(ws):
    paint(ws, 16, 28)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 36, "O": 13, "P": 3})
    for cl in month_cols():
        ws.column_dimensions[cl].width = 11
    banner(ws, "O", "🏦 PLAN DE TRÉSORERIE",
           "En clair : combien il y a réellement sur le compte, mois par "
           "mois — le nerf de la guerre.")
    tip(ws, "B3:O3",
        "Seules les lignes « Autres encaissements » et « Autres "
        "décaissements » se saisissent ici. Le reste provient de vos "
        "recettes encaissées, dépenses payées, salaire, cotisations, TVA, "
        "IS et CFE. Solde négatif = fond rouge, solde < 1 000 € = fond "
        "orange.")

    cols = month_cols()
    head(ws, "B5", "", TAUPE)
    for i, cl in enumerate(cols):
        head(ws, f"{cl}5", MOIS[i], TAUPE)
    head(ws, "O5", "Année", TAUPE)

    label(ws, "B6", "Solde en début de mois", bold=True)
    band(ws, "B7:O7", "ENCAISSEMENTS", SAUGE)
    label(ws, "B8", "Recettes encaissées (TTC)")
    label(ws, "B9", "Autres encaissements (apports, subventions…)")
    label(ws, "B10", "Total encaissements", bold=True)
    band(ws, "B11:O11", "DÉCAISSEMENTS", TERRA)
    label(ws, "B12", "Dépenses payées (TTC)")
    label(ws, "B13", "Salaire net du dirigeant")
    label(ws, "B14", "Cotisations sociales URSSAF (estimation)")
    label(ws, "B15", "TVA décaissée (estimation)")
    label(ws, "B16", "Acompte d'IS (estimation)")
    label(ws, "B17", "CFE (estimation)")
    label(ws, "B18", "Autres décaissements")
    label(ws, "B19", "Total décaissements", bold=True)
    label(ws, "B20", "SOLDE EN FIN DE MOIS", bold=True)

    for m in range(1, 13):
        cl = cols[m - 1]
        if m == 1:
            ccalc(ws, f"{cl}6", "=SOLDE_INIT", MONEY0, bold=True)
        else:
            ccalc(ws, f"{cl}6", f"={cols[m - 2]}20", MONEY0, bold=True)
        ccalc(ws, f"{cl}8",
              f"=SUMIFS('{REC}'!$H${REC_FIRST}:$H${REC_LAST},"
              f"'{REC}'!$J${REC_FIRST}:$J${REC_LAST},{m},"
              f"'{REC}'!$I${REC_FIRST}:$I${REC_LAST},\"Oui\")", MONEY0)
        cin(ws, f"{cl}9", 0, MONEY0)
        ccalc(ws, f"{cl}10", f"=SUM({cl}8:{cl}9)", MONEY0, bold=True)
        ccalc(ws, f"{cl}12",
              f"=SUMIFS('{DEP}'!$I${DEP_FIRST}:$I${DEP_LAST},"
              f"'{DEP}'!$K${DEP_FIRST}:$K${DEP_LAST},{m},"
              f"'{DEP}'!$J${DEP_FIRST}:$J${DEP_LAST},\"Oui\")", MONEY0)
        ccalc(ws, f"{cl}13", "=SALAIRE_NET", MONEY0)
        ccalc(ws, f"{cl}14", "=ROUND(SALAIRE_NET*TAUX_CHARGES,2)", MONEY0)
        ccalc(ws, f"{cl}15", f"=MAX(0,'{TVA}'!$E${6 + m})", MONEY0)
        ccalc(ws, f"{cl}16",
              f"=IF(OR({m}=3,{m}=6,{m}=9,{m}=12),'{TVA}'!$C$31,0)", MONEY0)
        ccalc(ws, f"{cl}17", f"=IF({m}=12,CFE_ANNUELLE,0)", MONEY0)
        cin(ws, f"{cl}18", 0, MONEY0)
        ccalc(ws, f"{cl}19", f"=SUM({cl}12:{cl}18)", MONEY0, bold=True)
        ccalc(ws, f"{cl}20", f"={cl}6+{cl}10-{cl}19", MONEY0, bold=True,
              font=F_TOTAL)

    ccalc(ws, "O6", "=C6", MONEY0, bold=True)
    for rr in (8, 9, 10, 12, 13, 14, 15, 16, 17, 18, 19):
        ccalc(ws, f"O{rr}", f"=SUM(C{rr}:N{rr})", MONEY0,
              bold=rr in (10, 19))
    ccalc(ws, "O20", "=N20", MONEY0, bold=True, font=F_TOTAL)

    ws.conditional_formatting.add("C20:O20", CellIsRule(
        operator="lessThan", formula=["0"], fill=fill(RED_L),
        font=Font(name=FONT, size=10, bold=True, color=TERRA),
        stopIfTrue=True))
    ws.conditional_formatting.add("C20:O20", CellIsRule(
        operator="lessThan", formula=["1000"], fill=fill(ORANGE_L)))
    ws.freeze_panes = "C6"


# ============================================================= TVA & IMPÔTS ==

def build_tva(ws):
    paint(ws, 11, 36)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 30, "C": 15, "D": 15, "E": 15, "F": 3,
                "G": 26, "H": 16, "I": 3})
    banner(ws, "H", "🧾 TVA & IMPÔTS",
           "En clair : ce que vous devrez reverser (TVA, URSSAF, IS, CFE) "
           "— estimé d'avance pour éviter les surprises.")
    tip(ws, "B3:H3",
        "Aucune saisie ici (sauf la base IS, modifiable). En franchise en "
        "base de TVA, les montants de TVA sont automatiquement à zéro. "
        "Montants indicatifs : vérifiez les échéances réelles avec votre "
        "comptable.")

    # --- TVA mensuelle ----------------------------------------------------
    band(ws, "B5:E5", "TVA — SYNTHÈSE MENSUELLE (ESTIMATION)")
    head(ws, "B6", "Mois", TAUPE)
    head(ws, "C6", "TVA collectée", TAUPE)
    head(ws, "D6", "TVA déductible", TAUPE)
    head(ws, "E6", "TVA à décaisser", TAUPE)
    for m in range(1, 13):
        r = 6 + m
        c = ws[f"B{r}"]
        c.value = MOIS[m - 1]
        c.font = F_LABEL_B
        c.fill = fill(BEIGE_ALT)
        c.border = BORDER
        c.alignment = A_C
        alt = BEIGE_ALT if m % 2 else BEIGE
        ccalc(ws, f"C{r}",
              f"=IF(REGIME_TVA=\"Franchise en base\",0,"
              f"SUMIFS('{REC}'!$G${REC_FIRST}:$G${REC_LAST},"
              f"'{REC}'!$J${REC_FIRST}:$J${REC_LAST},{m}))", MONEY, bg=alt)
        ccalc(ws, f"D{r}",
              f"=IF(REGIME_TVA=\"Franchise en base\",0,"
              f"SUMIFS('{DEP}'!$H${DEP_FIRST}:$H${DEP_LAST},"
              f"'{DEP}'!$K${DEP_FIRST}:$K${DEP_LAST},{m}))", MONEY, bg=alt)
        ccalc(ws, f"E{r}", f"=C{r}-D{r}", MONEY, bg=alt)
    label(ws, "B19", "Total annuel", bold=True)
    ws["B19"].font = F_TOTAL
    for cl in "CDE":
        ccalc(ws, f"{cl}19", f"=SUM({cl}7:{cl}18)", MONEY, bold=True,
              font=F_TOTAL)

    # --- TVA par trimestre ------------------------------------------------
    band(ws, "G5:H5", "TVA PAR TRIMESTRE (ESTIMATION)", TAUPE)
    for i, (lab, a, b) in enumerate([("T1 (janv.–mars)", 7, 9),
                                     ("T2 (avr.–juin)", 10, 12),
                                     ("T3 (juil.–sept.)", 13, 15),
                                     ("T4 (oct.–déc.)", 16, 18)]):
        r = 6 + i
        label(ws, f"G{r}", lab)
        ccalc(ws, f"H{r}", f"=SUM($E${a}:$E${b})", MONEY)
    merge(ws, "G10:H10", "Réel simplifié : acomptes en juillet et décembre, "
          "régularisation en mai.", F_NOTE, BEIGE, A_LW)

    # --- à provisionner ce mois-ci -----------------------------------------
    band(ws, "G12:H12", "💡 À PROVISIONNER CE MOIS-CI (ESTIMATION)")
    label(ws, "G13", "TVA du mois (estimation)")
    ccalc(ws, "H13", f"=MAX(0,INDEX($E$7:$E$18,{MIDX}))", MONEY)
    label(ws, "G14", "Cotisations URSSAF (estimation)")
    ccalc(ws, "H14", "=ROUND(SALAIRE_NET*TAUX_CHARGES,2)", MONEY)
    label(ws, "G15", "Acompte d'IS (estimation)")
    ccalc(ws, "H15",
          f"=IF(OR({MIDX}=3,{MIDX}=6,{MIDX}=9,{MIDX}=12),$C$31,0)", MONEY)
    label(ws, "G16", "CFE — 15 décembre (estimation)")
    ccalc(ws, "H16", f"=IF({MIDX}=12,CFE_ANNUELLE,0)", MONEY)
    label(ws, "G17", "TOTAL À PROVISIONNER", bold=True)
    ws["G17"].font = F_TOTAL
    ccalc(ws, "H17", "=SUM(H13:H16)", MONEY, bold=True, font=F_TOTAL)
    box(ws, 12, 7, 17, 8, BORDEAUX)

    # --- URSSAF -------------------------------------------------------------
    band(ws, "B21:E21",
         "COTISATIONS SOCIALES DU DIRIGEANT — URSSAF (ESTIMATION)")
    label(ws, "B22", "Statut")
    ccalc(ws, "C22", "=STATUT", "@").alignment = A_L
    merge(ws, "C22:E22")
    label(ws, "B23", "Taux de charges (sur le net)")
    ccalc(ws, "C23", "=TAUX_CHARGES", PCT0)
    label(ws, "B24", "Salaire net mensuel")
    ccalc(ws, "C24", "=SALAIRE_NET", MONEY)
    label(ws, "B25", "Cotisations mensuelles (estimation)")
    ccalc(ws, "C25", "=ROUND(SALAIRE_NET*TAUX_CHARGES,2)", MONEY, bold=True)
    label(ws, "B26", "Cotisations annuelles (estimation)")
    ccalc(ws, "C26", "=$C$25*12", MONEY, bold=True)
    merge(ws, "B27:E27",
          "Taux moyens indicatifs (modifiables dans ⚙️ Paramètres), à "
          "ajuster selon votre situation.", F_NOTE, BEIGE, A_LW)

    # --- IS -------------------------------------------------------------------
    band(ws, "B28:E28", "IMPÔT SUR LES SOCIÉTÉS (ESTIMATION)")
    label(ws, "B29", "Base : résultat prévisionnel annuel (modifiable)")
    cin(ws, "C29", f"='{BUD}'!$O$33", MONEY0)
    label(ws, "B30", "IS estimé (barème 15 % / 25 %)")
    ccalc(ws, "C30",
          "=IF($C$29<=0,0,IF($C$29<=SEUIL_IS,ROUND($C$29*TAUX_IS_REDUIT,2),"
          "ROUND(SEUIL_IS*TAUX_IS_REDUIT+($C$29-SEUIL_IS)*TAUX_IS_NORMAL,2)))",
          MONEY, bold=True)
    label(ws, "B31", "Acompte trimestriel (IS ÷ 4)")
    ccalc(ws, "C31", "=ROUND($C$30/4,2)", MONEY)
    merge(ws, "B32:E32", "Échéances des acomptes : 15 mars · 15 juin · "
          "15 septembre · 15 décembre.", F_NOTE, BEIGE, A_LW)

    # --- CFE --------------------------------------------------------------------
    band(ws, "G21:H21", "CFE (ESTIMATION)", TAUPE)
    label(ws, "G22", "Montant estimé (⚙️ Paramètres)")
    ccalc(ws, "H22", "=CFE_ANNUELLE", MONEY)
    label(ws, "G23", "Échéance")
    c = ws["H23"]
    c.value = "15 décembre"
    c.font = F_CALC_B
    c.fill = fill(BEIGE_ALT)
    c.border = BORDER
    c.alignment = A_C


# ========================================================== INDEMNITÉS KM ==

KM_SAMPLES = [
    (dt.date(2026, 1, 8), "RDV client — Studio Novelli", "Vannes", "Nantes",
     220),
    (dt.date(2026, 1, 15), "Signature contrat", "Vannes", "Lorient", 110),
    (dt.date(2026, 1, 22), "Formation — Mairie de Beaulieu", "Vannes",
     "Beaulieu", 86),
    (dt.date(2026, 1, 29), "Prospection réseau d'affaires", "Vannes",
     "Rennes", 230),
    (dt.date(2026, 2, 5), "Atelier — Éditions Clairval", "Vannes", "Auray",
     42),
    (dt.date(2026, 2, 12), "Suivi mensuel — Studio Novelli", "Vannes",
     "Nantes", 220),
    (dt.date(2026, 2, 19), "Salon des entrepreneurs", "Vannes", "Rennes",
     230),
    (dt.date(2026, 3, 4), "Diagnostic — Groupe Vaillant", "Vannes", "Brest",
     370),
    (dt.date(2026, 3, 12), "RDV expert-comptable", "Vannes", "Vannes centre",
     12),
    (dt.date(2026, 3, 18), "Formation — Mairie de Beaulieu", "Vannes",
     "Beaulieu", 86),
]


def build_km(ws):
    paint(ws, 10, KM_LAST + 6)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 12, "C": 36, "D": 18, "E": 18, "F": 12,
                "G": 3, "H": 28, "I": 16, "J": 3})
    banner(ws, "I", "🚗 INDEMNITÉS KILOMÉTRIQUES",
           "En clair : vos trajets pro deviennent un remboursement "
           "défiscalisé — notez-les, c'est de l'argent.")
    tip(ws, "B3:I3",
        "Choisissez la puissance fiscale de votre véhicule, puis saisissez "
        "chaque déplacement (km aller-retour). L'indemnité est calculée sur "
        "le total annuel selon le barème de l'onglet ⚙️ Paramètres "
        "(tranches ≤ 5 000 km / 5 001–20 000 km / au-delà).")

    band(ws, "H5:I5", "VOTRE VÉHICULE & VOS INDEMNITÉS")
    label(ws, "H6", "Puissance fiscale")
    cin(ws, "I6", "5 CV")
    label(ws, "H7", "Total km sur l'année")
    ccalc(ws, "I7", f"=SUM($F${KM_FIRST}:$F${KM_LAST})", KMF, bold=True)
    label(ws, "H8", "Indemnité kilométrique estimée")
    ccalc(ws, "I8",
          "=IFERROR(IF($I$7<=0,0,IF($I$7<=5000,"
          "ROUND($I$7*VLOOKUP($I$6,BAREME_KM,2,0),2),"
          "IF($I$7<=20000,"
          "ROUND($I$7*VLOOKUP($I$6,BAREME_KM,3,0)+VLOOKUP($I$6,BAREME_KM,4,0),2),"
          "ROUND($I$7*VLOOKUP($I$6,BAREME_KM,5,0),2)))),0)",
          MONEY, bold=True, font=F_TOTAL)
    box(ws, 5, 8, 8, 9, BORDEAUX)

    headers = ["Date", "Motif du déplacement", "Départ", "Arrivée",
               "Km A/R"]
    for i, h in enumerate(headers):
        head(ws, f"{get_column_letter(2 + i)}9", h)
    ws.row_dimensions[9].height = 22

    for r in range(KM_FIRST, KM_LAST + 1):
        cin(ws, f"B{r}", fmt=DATEF)
        cin(ws, f"C{r}").alignment = A_L
        cin(ws, f"D{r}").alignment = A_L
        cin(ws, f"E{r}").alignment = A_L
        cin(ws, f"F{r}", fmt=KMF)

    for i, s in enumerate(KM_SAMPLES):
        r = KM_FIRST + i
        ws[f"B{r}"] = s[0]
        ws[f"C{r}"] = s[1]
        ws[f"D{r}"] = s[2]
        ws[f"E{r}"] = s[3]
        ws[f"F{r}"] = s[4]

    dv_cv = DataValidation(
        type="list",
        formula1='"3 CV et moins,4 CV,5 CV,6 CV,7 CV et plus"',
        allow_blank=True)
    ws.add_data_validation(dv_cv)
    dv_cv.add("I6")
    ws.freeze_panes = "A10"


# ================================================================ DASHBOARD ==

def kpi_card(ws, col1, col2, r, titre, formule, fmt, sous_titre):
    """Carte KPI premium : bandeau bordeaux, gros chiffre, sous-titre."""
    merge(ws, f"{col1}{r}:{col2}{r}", titre.upper(),
          Font(name=FONT, size=8, bold=True, color="FFFFFF"), BORDEAUX,
          A_C)
    v = merge(ws, f"{col1}{r + 1}:{col2}{r + 1}", formule,
              Font(name=FONT, size=16, bold=True, color=BORDEAUX), CREME,
              A_C)
    v.number_format = fmt
    merge(ws, f"{col1}{r + 2}:{col2}{r + 2}", sous_titre, F_KPI_SUB, CREME,
          A_CW)
    c1 = ws[f"{col1}{r}"].column
    c2 = ws[f"{col2}{r}"].column
    box(ws, r, c1, r + 2, c2, TAUPE)


def build_dashboard(ws):
    paint(ws, 17, 66)
    ws.sheet_properties.tabColor = BORDEAUX
    widths(ws, {"A": 2.5, "N": 3, "P": 30, "Q": 14})
    for cl in "BCDEFGHIJKLM":
        ws.column_dimensions[cl].width = 12.5
    banner(ws, "N", '="🏠  "&UPPER(NOM_ENT)&" — TABLEAU DE BORD "&ANNEE',
           "En clair : toute votre entreprise en un coup d'œil — santé, "
           "alertes et chiffres clés.")

    # rappel de la configuration (onglet 🚀 Démarrer ici) + zone 1
    c = ws["B3"]
    c.value = "V U E   D ' E N S E M B L E"
    c.font = Font(name=FONT, size=9, bold=True, color=TAUPE)
    c.alignment = A_L
    merge(ws, "K3:M3",
          f'="🚀 Configuration : "&COUNTIF(\'{START}\'!$I$7:$I$19,"✅")'
          f'&"/5 étapes"',
          Font(name=FONT, size=8.5, italic=True, color=TAUPE), BEIGE,
          Alignment(horizontal="right", vertical="center"))

    # --- cartes KPI --------------------------------------------------------
    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 26
    ws.row_dimensions[6].height = 22
    kpi_card(ws, "B", "C", 4, "CA réalisé (HT)",
             f"='{SUI}'!$D$19", MONEY0, "chiffre d'affaires facturé")
    kpi_card(ws, "D", "E", 4, "Dépenses réalisées",
             f"='{SUI}'!$G$19", MONEY0, "HT + salaire & charges")
    kpi_card(ws, "F", "G", 4, "Résultat",
             f"='{SUI}'!$J$19", MONEY0, "CA − dépenses")
    kpi_card(ws, "H", "I", 4, "Trésorerie actuelle",
             f"=INDEX('{TRE}'!$C$20:$N$20,{MIDX})", MONEY0,
             "solde fin du mois en cours")
    kpi_card(ws, "J", "K", 4, "TVA estimée à payer",
             f"='{TVA}'!$H$13", MONEY0, "estimation — mois en cours")
    kpi_card(ws, "L", "M", 4, "Prochaine échéance",
             f"=IF('{ECH}'!$J$5=0,\"—\",'{ECH}'!$J$5)", DATEF,
             f"=IF('{ECH}'!$J$5=0,\"aucune échéance à venir\","
             f"'{ECH}'!$J$6&\" · \"&ROUND('{ECH}'!$J$7,0)&\" € · dans \""
             f"&'{ECH}'!$J$8&\" j\")")
    ws.conditional_formatting.add("F5", CellIsRule(
        operator="lessThan", formula=["0"],
        font=Font(name=FONT, size=16, bold=True, color=TERRA)))
    ws.conditional_formatting.add("H5", CellIsRule(
        operator="lessThan", formula=["0"],
        font=Font(name=FONT, size=16, bold=True, color=TERRA)))

    # --- zone de calcul (colonnes P/Q, hors écran principal) ----------------
    tiny = Font(name=FONT, size=8, color=TAUPE)
    merge(ws, "P3:Q3", "Zone de calcul — ne pas supprimer", tiny, BEIGE, A_L)
    helpers = [
        (4, "Mois en cours (borné à l'année)", f"={MIDX}", "0"),
        (5, "CA réalisé cumulé à date",
         f"=SUM('{SUI}'!$D$7:INDEX('{SUI}'!$D$7:$D$18,$Q$4))", MONEY0),
        (6, "Dépenses réalisées cumulées à date",
         f"=SUM('{SUI}'!$G$7:INDEX('{SUI}'!$G$7:$G$18,$Q$4))", MONEY0),
        (7, "CA prévu cumulé à date",
         f"=SUM('{SUI}'!$C$7:INDEX('{SUI}'!$C$7:$C$18,$Q$4))", MONEY0),
        (8, "Dépenses prévues cumulées à date",
         f"=SUM('{SUI}'!$F$7:INDEX('{SUI}'!$F$7:$F$18,$Q$4))", MONEY0),
        (9, "Trésorerie actuelle",
         f"=INDEX('{TRE}'!$C$20:$N$20,$Q$4)", MONEY0),
        (10, "Décaissements mensuels moyens",
         f"=IFERROR(SUM('{TRE}'!$C$19:INDEX('{TRE}'!$C$19:$N$19,$Q$4))"
         f"/$Q$4,0)", MONEY0),
        (11, "Runway (mois d'autonomie)",
         "=IFERROR(IF($Q$10<=0,99,MAX(0,$Q$9)/$Q$10),0)", "0.0"),
        (12, "Marge nette réalisée",
         "=IFERROR(($Q$5-$Q$6)/$Q$5,0)", PCT),
        (13, "Dépenses réalisées / prévues à date",
         "=IFERROR($Q$6/$Q$8,1)", PCT),
        (14, "CA réalisé / prévu à date",
         "=IFERROR($Q$5/$Q$7,0)", PCT),
        (15, "Provisions du mois (estimation)",
         f"='{TVA}'!$H$17", MONEY0),
        (16, "Factures non encaissées (TTC)",
         f"=SUMIFS('{REC}'!$H${REC_FIRST}:$H${REC_LAST},"
         f"'{REC}'!$I${REC_FIRST}:$I${REC_LAST},\"Non\")", MONEY0),
        (17, "Seuil de rentabilité mensuel",
         "=IFERROR($Q$6/$Q$4,0)", MONEY0),
        (18, "Résultat projeté fin d'année (run-rate)",
         "=IFERROR(($Q$5-$Q$6)/$Q$4*12,0)", MONEY0),
        (19, "Données saisies ? (1 = oui)",
         f"=IF($Q$5+'{DEP}'!$O$6=0,0,1)", "0"),
        (20, "Score trésorerie /100",
         "=MIN(100,MAX(0,ROUND($Q$11/6*100,0)))", "0"),
        (21, "Score rentabilité /100",
         "=MIN(100,MAX(0,ROUND($Q$12/0.2*100,0)))", "0"),
        (22, "Score respect du budget /100",
         "=MIN(100,MAX(0,ROUND(100-MAX(0,$Q$13-1)*200,0)))", "0"),
        (23, "Score dynamique CA /100",
         "=MIN(100,MAX(0,ROUND($Q$14*100,0)))", "0"),
        (24, "Score provisions /100",
         "=IF($Q$9>=$Q$15,100,0)", "0"),
        (25, "SCORE GLOBAL /100",
         "=ROUND(0.3*$Q$20+0.25*$Q$21+0.2*$Q$22+0.15*$Q$23+0.1*$Q$24,0)",
         "0"),
        (26, "CA facturé TTC (année)",
         f"=SUM('{REC}'!$H${REC_FIRST}:$H${REC_LAST})", MONEY0),
    ]
    # dépenses réalisées par catégorie du budget (pour le Top 5) ; le
    # micro-ajout (ROW()/1e6) départage les ex æquo sans fausser l'affichage
    for i in range(15):
        r = 28 + i
        helpers.append((r, f"='{BUD}'!$B${16 + i}",
                        f"=SUMIFS('{DEP}'!$F${DEP_FIRST}:$F${DEP_LAST},"
                        f"'{DEP}'!$D${DEP_FIRST}:$D${DEP_LAST},$P{r})"
                        f"+(ROW()-27)/1000000", MONEY0))
    for r, lab, formula, *fmt in helpers:
        lc = ws[f"P{r}"]
        lc.value = lab
        lc.font = tiny
        lc.alignment = A_L
        vc = ws[f"Q{r}"]
        vc.value = formula
        vc.font = tiny
        vc.alignment = A_R
        vc.number_format = fmt[0] if fmt else NUMF

    # --- indicateurs compacts (sous les cartes KPI) ---------------------------
    ws.row_dimensions[8].height = 13
    ws.row_dimensions[9].height = 22
    chips = [
        ("B", "D", "⏳ RUNWAY",
         '=IF($Q$19=0,"—",ROUND($Q$11,1)&" mois d\'autonomie")'),
        ("E", "G", "📈 MARGE NETTE",
         '=IF($Q$19=0,"—",ROUND($Q$12*100,0)&" % du CA")'),
        ("H", "J", "🎯 SEUIL DE RENTABILITÉ",
         '=IF($Q$19=0,"—",ROUND($Q$17,0)&" €/mois de CA pour couvrir '
         'vos charges")'),
        ("K", "M", "🔮 RÉSULTAT PROJETÉ",
         '=IF($Q$19=0,"—",ROUND($Q$18,0)&" € fin d\'année (run-rate)")'),
    ]
    for c1, c2, lab, formula in chips:
        merge(ws, f"{c1}8:{c2}8", lab, F_KPI_LAB, CREME, A_C)
        merge(ws, f"{c1}9:{c2}9", formula,
              Font(name=FONT, size=9.5, bold=True, color=BORDEAUX), CREME,
              A_CW)
        col1 = ws[f"{c1}8"].column
        col2 = ws[f"{c2}8"].column
        box(ws, 8, col1, 9, col2, TAUPE_L)

    # --- zone 2 : santé & recommandations -----------------------------------
    band(ws, "B11:M11", "S A N T É   &   R E C O M M A N D A T I O N S",
         TAUPE)
    merge(ws, "B12:C12", "🩺 SANTÉ FINANCIÈRE", F_KPI_LAB, CREME, A_C)
    merge(ws, "D12:G12",
          '=IF($Q$19=0,"— en attente de données",IF($Q$25>=75,'
          '"🟢 Solide",IF($Q$25>=40,"🟡 À surveiller","🔴 Fragile")))',
          Font(name=FONT, size=11, bold=True, color=INK), CREME, A_L)
    merge(ws, "B13:C13", '=IF($Q$19=0,"—",$Q$25&" / 100")',
          Font(name=FONT, size=18, bold=True, color=BORDEAUX), CREME, A_C)
    merge(ws, "D13:G13",
          '=IF($Q$19=0,"",REPT("█",ROUND($Q$25/5,0))'
          '&REPT("░",20-ROUND($Q$25/5,0)))',
          Font(name=FONT, size=11, color=BORDEAUX), CREME, A_L)
    ws.row_dimensions[13].height = 24
    composantes = [
        (14, "Trésorerie (30 %)", "$Q$20",
         '"Autonomie : "&ROUND($Q$11,1)&" mois de décaissements devant '
         'vous (objectif : 6)."'),
        (15, "Rentabilité (25 %)", "$Q$21",
         '"Marge nette : "&ROUND($Q$12*100,0)&" % du CA (objectif : '
         '20 %)."'),
        (16, "Respect du budget (20 %)", "$Q$22",
         '"Dépenses : "&ROUND($Q$13*100,0)&" % du budget prévu à date '
         '(objectif : 100 % maximum)."'),
        (17, "Dynamique CA (15 %)", "$Q$23",
         '"CA réalisé : "&ROUND($Q$14*100,0)&" % du prévisionnel à '
         'date."'),
        (18, "Provisions (10 %)", "$Q$24",
         'IF($Q$24=100,"Les provisions du mois sont couvertes par la '
         'trésorerie.","La trésorerie ne couvre pas encore les provisions '
         'du mois.")'),
    ]
    for r, lab, score_ref, phrase in composantes:
        merge(ws, f"B{r}:C{r}", lab, F_LABEL_B, CREME, A_L)
        sc = ws[f"D{r}"]
        sc.value = f'=IF($Q$19=0,"—",{score_ref}&"/100")'
        sc.font = F_CALC_B
        sc.fill = fill(CREME)
        sc.alignment = A_C
        merge(ws, f"E{r}:G{r}",
              f'=IF($Q$19=0,"— en attente de données",{phrase})',
              F_NOTE, CREME, A_LW)
        ws.row_dimensions[r].height = 28
    ws.row_dimensions[18].height = 18
    box(ws, 12, 2, 18, 7, TAUPE)

    # --- recommandations contextuelles ---------------------------------------
    band(ws, "H12:M12", "🧭 VOS RECOMMANDATIONS DU MOMENT")
    attente = "— En attente de vos premières saisies."
    recos = [
        (13,
         f'=IF($Q$19=0,"{attente}",IF($Q$11<3,'
         '"⚠️ Moins de 3 mois d\'autonomie de trésorerie. Priorité : '
         'relancez vos factures non encaissées ("&ROUND($Q$16,0)&" € en '
         'attente) et décalez les dépenses non urgentes.",'
         '"✓ RAS — trésorerie : "&ROUND($Q$11,1)&" mois d\'autonomie '
         'devant vous."))'),
        (14,
         f'=IF($Q$19=0,"{attente}",IF($Q$13>1.1,'
         '"📊 Vos dépenses dépassent le budget de "&ROUND(($Q$13-1)*100,0)'
         '&" %. Consultez le 📊 Suivi mensuel pour repérer la catégorie '
         'en cause.",'
         '"✓ RAS — budget respecté à ce stade de l\'année."))'),
        (15,
         f"=IF('{ECH}'!$J$5=0,\"✓ RAS — aucune échéance à venir.\","
         f"IF(AND('{ECH}'!$J$8<15,$Q$9<'{ECH}'!$J$7),"
         f"\"📅 Échéance \"&'{ECH}'!$J$6&\" dans \"&'{ECH}'!$J$8"
         f"&\" j : prévoyez \"&ROUND('{ECH}'!$J$7,0)"
         f'&" € sur le compte.",'
         f'"✓ RAS — prochaine échéance sous contrôle."))'),
        (16,
         f'=IF($Q$19=0,"{attente}",IF(AND($Q$26>0,$Q$16>0.2*$Q$26),'
         '"💶 "&ROUND($Q$16,0)&" € facturés mais non encaissés. Pensez '
         'aux relances — c\'est de la trésorerie qui dort.",'
         '"✓ RAS — encaissements à jour."))'),
        (17,
         '=IF($Q$19=0,"🌱 Commencez à saisir vos factures et vos dépenses '
         ': vos recommandations personnalisées s\'afficheront ici.",'
         'IF(COUNTIF($H$13:$H$16,"✓*")=4,'
         '"🎉 Tous les voyants sont au vert ce mois-ci. Continuez comme '
         'ça !",'
         '"☝️ Traitez les points ci-dessus — le reste est sous '
         'contrôle."))'),
    ]
    for r, formula in recos:
        merge(ws, f"H{r}:M{r}", formula,
              Font(name=FONT, size=9, color=INK), CREME, A_LW)
    merge(ws, "H18:M18",
          "Recommandations générales, à titre indicatif.",
          Font(name=FONT, size=8, italic=True, color=TAUPE), CREME, A_L)
    box(ws, 12, 8, 18, 13, TAUPE)

    # --- zone 3 : analyses ----------------------------------------------------
    band(ws, "B20:M20", "A N A L Y S E S", TAUPE)

    def style_chart(chart, title, gridlines=True):
        """Types natifs simples, fond beige sans bordure, titre bordeaux,
        quadrillage discret — importables tels quels dans Google Sheets."""
        chart.title = title
        try:
            para = chart.title.tx.rich.p[0]
            cp = CharacterProperties(sz=1150, b=True, solidFill=BORDEAUX)
            para.pPr.defRPr = cp
            for run in para.r:
                run.rPr = cp
        except (AttributeError, IndexError, TypeError):
            pass
        chart.graphical_properties = GraphicalProperties(
            solidFill=BEIGE, ln=LineProperties(noFill=True))
        if gridlines:
            chart.y_axis.majorGridlines = ChartLines(
                spPr=GraphicalProperties(
                    ln=LineProperties(solidFill=TAUPE_L, w=9525)))

    bar = BarChart()
    bar.type = "col"
    bar.grouping = "clustered"
    style_chart(bar, "Recettes vs Dépenses (réalisé)")
    bar.gapWidth = 60
    cats = Reference(ws.parent[SUI], min_col=2, min_row=7, max_row=18)
    s1 = Series(Reference(ws.parent[SUI], min_col=4, min_row=7, max_row=18),
                title="Recettes")
    s2 = Series(Reference(ws.parent[SUI], min_col=7, min_row=7, max_row=18),
                title="Dépenses")
    s1.graphicalProperties.solidFill = BORDEAUX
    s2.graphicalProperties.solidFill = TAUPE
    bar.append(s1)
    bar.append(s2)
    bar.set_categories(cats)
    bar.legend.position = "b"
    bar.y_axis.numFmt = '#,##0 "€"'
    bar.x_axis.delete = False
    bar.y_axis.delete = False
    bar.width = 13.2
    bar.height = 8.2
    ws.add_chart(bar, "B21")

    line = LineChart()
    line.grouping = "standard"
    style_chart(line, "Trésorerie — solde de fin de mois")
    tre = ws.parent[TRE]
    s3 = Series(Reference(tre, min_col=3, max_col=14, min_row=20,
                          max_row=20), title="Trésorerie")
    s3.graphicalProperties.line.solidFill = BORDEAUX
    s3.graphicalProperties.line.width = 28575
    s3.smooth = True
    s3.marker = Marker(symbol="circle", size=5,
                       spPr=GraphicalProperties(
                           solidFill=TERRA,
                           ln=LineProperties(solidFill=CREME, w=9525)))
    line.append(s3)
    line.set_categories(Reference(tre, min_col=3, max_col=14, min_row=5,
                                  max_row=5))
    line.legend = None
    line.y_axis.numFmt = '#,##0 "€"'
    line.x_axis.delete = False
    line.y_axis.delete = False
    line.width = 13.2
    line.height = 8.2
    ws.add_chart(line, "H21")

    # --- mini-tableau prévisionnel vs réalisé -------------------------------
    band(ws, "B38:H38", "PRÉVISIONNEL vs RÉALISÉ — CUMUL ANNUEL")
    head(ws, "B39", "", TAUPE)
    head(ws, "C39", "Prévu", TAUPE)
    head(ws, "D39", "Réalisé", TAUPE)
    head(ws, "E39", "Écart", TAUPE)
    head(ws, "F39", "Écart %", TAUPE)
    head(ws, "G39", "Tendance", TAUPE)
    head(ws, "H39", "Progression", TAUPE)
    rows = [
        ("Recettes", f"='{BUD}'!$O$12", f"='{SUI}'!$D$19",
         "=D40-C40", "=IFERROR(E40/C40,0)"),
        ("Dépenses", f"='{BUD}'!$O$31", f"='{SUI}'!$G$19",
         "=C41-D41", "=IFERROR(E41/C41,0)"),
        ("Résultat", f"='{BUD}'!$O$33", f"='{SUI}'!$J$19",
         "=D42-C42", "=IFERROR(E42/ABS(C42),0)"),
    ]
    for i, (lab, prev, real, ecart, pc) in enumerate(rows):
        r = 40 + i
        c = ws[f"B{r}"]
        c.value = lab
        c.font = F_LABEL_B
        c.fill = fill(BEIGE_ALT)
        c.border = BORDER
        c.alignment = A_L
        ccalc(ws, f"C{r}", prev, MONEY0)
        ccalc(ws, f"D{r}", real, MONEY0)
        ccalc(ws, f"E{r}", ecart, MONEY0)
        ccalc(ws, f"F{r}", pc, PCT)
        fl = ccalc(ws, f"G{r}", f'=IF(E{r}>=0,"▲","▼")', "@")
        fl.alignment = A_C
        # mini-barre REPT() : part du réalisé par rapport au prévu
        ratio = f"MIN(1,MAX(0,IFERROR(D{r}/C{r},0)))"
        bar_cell = ws[f"H{r}"]
        bar_cell.value = (f'=REPT("█",ROUND({ratio}*10,0))'
                          f'&REPT("░",10-ROUND({ratio}*10,0))')
        bar_cell.font = Font(name=FONT, size=9, color=BORDEAUX)
        bar_cell.fill = fill(CREME)
        bar_cell.border = BORDER
        bar_cell.alignment = A_L
    sauge_f = Font(name=FONT, size=10, bold=True, color=SAUGE)
    terra_f = Font(name=FONT, size=10, bold=True, color=TERRA)
    for cl in ("E", "G"):
        ws.conditional_formatting.add(f"{cl}40:{cl}42", FormulaRule(
            formula=["$E40>=0"], font=sauge_f))
        ws.conditional_formatting.add(f"{cl}40:{cl}42", FormulaRule(
            formula=["$E40<0"], font=terra_f))

    # --- top 5 des dépenses de l'année ----------------------------------------
    band(ws, "I38:M38", "TOP 5 DES DÉPENSES DE L'ANNÉE", TAUPE)
    merge(ws, "I39:J39", "Catégorie", F_HEAD, TAUPE, A_CW, BORDER)
    head(ws, "K39", "Montant", TAUPE)
    merge(ws, "L39:M39", "Poids", F_HEAD, TAUPE, A_CW, BORDER)
    for k in range(1, 6):
        r = 39 + k
        big = f"LARGE($Q$28:$Q$42,{k})"
        merge(ws, f"I{r}:J{r}",
              f'=IFERROR(IF({big}<0.005,"—",'
              f"INDEX($P$28:$P$42,MATCH({big},$Q$28:$Q$42,0))),\"—\")",
              F_LABEL, BEIGE_ALT, A_LW, BORDER)
        ccalc(ws, f"K{r}", f"=IFERROR(ROUND({big},0),0)", MONEY0)
        share = f"{big}/MAX($Q$28:$Q$42)"
        merge(ws, f"L{r}:M{r}",
              f'=IFERROR(IF(MAX($Q$28:$Q$42)<0.005,"",'
              f'REPT("█",ROUND({share}*10,0))'
              f'&REPT("░",10-ROUND({share}*10,0))),"")',
              Font(name=FONT, size=9, color=BORDEAUX), CREME, A_L, BORDER)

    # --- donut : répartition des 5 plus grosses dépenses ---------------------
    pie = DoughnutChart(holeSize=58)
    style_chart(pie, "Où part votre argent ? (Top 5)", gridlines=False)
    pie.add_data(Reference(ws, min_col=11, min_row=40, max_row=44),
                 titles_from_data=False)
    pie.set_categories(Reference(ws, min_col=9, min_row=40, max_row=44))
    pie.series[0].data_points = [
        DataPoint(idx=i, spPr=GraphicalProperties(
            solidFill=col, ln=LineProperties(solidFill=CREME, w=19050)))
        for i, col in enumerate([BORDEAUX, TERRA, TAUPE, SAUGE, TAUPE_L])
    ]
    pie.legend.position = "r"
    pie.width = 13.2
    pie.height = 7.4
    ws.add_chart(pie, "B46")

    # --- indemnités km + jauge budget ----------------------------------------
    merge(ws, "I46:M46",
          f'="🚗 Indemnités kilométriques : "&ROUND(\'{KM}\'!$I$7,0)'
          f'&" km · "&ROUND(\'{KM}\'!$I$8,0)&" € (estimation)"',
          F_LABEL_B, BEIGE_ALT, A_L, BORDER)
    label(ws, "I48", "Budget annuel consommé")
    ccalc(ws, "K48", f"=IFERROR('{SUI}'!$G$19/'{BUD}'!$O$31,0)", PCT0,
          bold=True)
    # barre de progression 100 % formule (compatible Excel + Google Sheets)
    gauge = "MIN(1,MAX(0,$K$48))"
    merge(ws, "L48:M48",
          f'=REPT("█",ROUND({gauge}*20,0))'
          f'&REPT("░",20-ROUND({gauge}*20,0))',
          Font(name=FONT, size=9, color=BORDEAUX), CREME, A_L, BORDER)
    merge(ws, "I50:M50",
          '="💶 En attente d\'encaissement : "&ROUND($Q$16,0)&" € — '
          'pensez aux relances."',
          F_NOTE, BEIGE, A_LW)

    # --- disclaimer -----------------------------------------------------------
    merge(ws, "B62:M62",
          "Outil de pilotage et d'estimation. Les montants de TVA, IS, CFE "
          "et cotisations sont indicatifs et ne remplacent pas l'avis d'un "
          "expert-comptable.", F_NOTE, BEIGE, A_CW)


# ============================================== stubs (phases suivantes) ===

def build_echeancier(ws):
    paint(ws, 11, 50)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 13, "C": 18, "D": 42, "E": 14, "F": 12,
                "G": 22, "H": 3, "I": 30, "J": 14})
    banner(ws, "J", "📅 ÉCHÉANCIER FISCAL & SOCIAL",
           "En clair : toutes vos dates limites (TVA, URSSAF, IS, CFE) au "
           "même endroit — fini les mauvaises surprises.")
    tip(ws, "B3:J3",
        "Les lignes s'adaptent automatiquement à votre régime de TVA "
        "(⚙️ Paramètres). Passez une échéance en « Payé » quand c'est "
        "réglé : le compte à rebours et les totaux se mettent à jour "
        "seuls.")

    first, last = 11, 43

    # --- zone de calcul -------------------------------------------------
    tiny = Font(name=FONT, size=8, color=TAUPE)
    merge(ws, "I4:J4", "Zone de calcul — ne pas supprimer", tiny, BEIGE,
          A_L)
    ech_helpers = [
        (5, "Prochaine échéance (date)",
         f"=IF(COUNT($H${first}:$H${last})=0,0,MIN($H${first}:$H${last}))",
         DATEF),
        (6, "Prochaine échéance (nature)",
         f"=IF($J$5=0,\"—\",INDEX($D${first}:$D${last},"
         f"MATCH($J$5,$B${first}:$B${last},0)))", "@"),
        (7, "Prochaine échéance (montant)",
         f"=IF($J$5=0,0,INDEX($E${first}:$E${last},"
         f"MATCH($J$5,$B${first}:$B${last},0)))", MONEY0),
        (8, "Prochaine échéance (jours)",
         "=IF($J$5=0,0,$J$5-TODAY())", "0"),
        (9, "À prévoir ce mois-ci",
         f"=SUMIFS($E${first}:$E${last},$B${first}:$B${last},"
         f'">="&DATE(ANNEE,{MIDX},1),$B${first}:$B${last},'
         f'"<"&DATE(ANNEE,{MIDX}+1,1),$F${first}:$F${last},"<>Payé")',
         MONEY0),
        (10, "Reste à payer sur l'année",
         f"=SUMIFS($E${first}:$E${last},$B${first}:$B${last},"
         f'">="&TODAY(),$F${first}:$F${last},"<>Payé")', MONEY0),
    ]
    for r, lab, formula, fmt in ech_helpers:
        lc = ws[f"I{r}"]
        lc.value = lab
        lc.font = tiny
        lc.alignment = A_L
        vc = ws[f"J{r}"]
        vc.value = formula
        vc.font = tiny
        vc.alignment = A_R
        vc.number_format = fmt

    # --- cartes ----------------------------------------------------------
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 26
    ws.row_dimensions[7].height = 20
    kpi_card(ws, "B", "C", 5, "Prochaine échéance",
             '=IF($J$5=0,"—",$J$5)', DATEF,
             '=IF($J$5=0,"aucune échéance à venir",$J$6&" · dans "&$J$8'
             '&" j")')
    kpi_card(ws, "D", "E", 5, "À prévoir ce mois-ci", "=$J$9", MONEY0,
             "échéances non payées du mois (estimation)")
    kpi_card(ws, "F", "G", 5, "Reste à payer sur l'année", "=$J$10",
             MONEY0, "toutes échéances à venir (estimation)")

    # --- tableau des échéances -------------------------------------------
    for i, h in enumerate(["Date", "Organisme", "Nature",
                           "Montant estimé", "Statut",
                           "Compte à rebours"]):
        head(ws, f"{get_column_letter(2 + i)}10", h)
    ws.row_dimensions[10].height = 22

    urssaf_amt = "=ROUND(SALAIRE_NET*TAUX_CHARGES,2)"
    rows = []
    for m in range(1, 13):
        rows.append((f"=DATE(ANNEE,{m},JOUR_URSSAF*1)", "URSSAF",
                     f"Cotisations sociales — {MOIS_FR[m - 1]} "
                     f"(estimation)", urssaf_amt))
        due = (f"DATE(ANNEE,{m + 1},15)" if m < 12
               else "DATE(ANNEE+1,1,15)")
        rows.append((f'=IF(REGIME_TVA="Réel normal",{due},"")',
                     "Impôts (TVA)",
                     f"TVA de {MOIS_FR[m - 1]} (estimation)",
                     f'=IF(REGIME_TVA="Réel normal",'
                     f"MAX(0,'{TVA}'!$E${6 + m}),\"\")"))
    rows += [
        ('=IF(REGIME_TVA="Réel simplifié",DATE(ANNEE,7,15),"")',
         "Impôts (TVA)",
         "Acompte de TVA de juillet — réel simplifié (estimation)",
         f'=IF(REGIME_TVA="Réel simplifié",'
         f"MAX(0,ROUND('{TVA}'!$E$19*TVA_AC_JUIL,2)),\"\")"),
        ('=IF(REGIME_TVA="Réel simplifié",DATE(ANNEE,12,15),"")',
         "Impôts (TVA)",
         "Acompte de TVA de décembre — réel simplifié (estimation)",
         f'=IF(REGIME_TVA="Réel simplifié",'
         f"MAX(0,ROUND('{TVA}'!$E$19*TVA_AC_DEC,2)),\"\")"),
        ('=IF(REGIME_TVA="Réel simplifié",DATE(ANNEE+1,5,3),"")',
         "Impôts (TVA)",
         "Solde de TVA (déclaration CA12) — mai N+1 (estimation)",
         f'=IF(REGIME_TVA="Réel simplifié",'
         f"MAX(0,ROUND('{TVA}'!$E$19*(1-TVA_AC_JUIL-TVA_AC_DEC),2)),"
         f'"")'),
        ("=DATE(ANNEE,3,15)", "Impôts (IS)",
         "1er acompte d'IS (estimation)", f"='{TVA}'!$C$31"),
        ("=DATE(ANNEE,6,15)", "Impôts (IS)",
         "2e acompte d'IS (estimation)", f"='{TVA}'!$C$31"),
        ("=DATE(ANNEE,9,15)", "Impôts (IS)",
         "3e acompte d'IS (estimation)", f"='{TVA}'!$C$31"),
        ("=DATE(ANNEE,12,15)", "Impôts (IS)",
         "4e acompte d'IS (estimation)", f"='{TVA}'!$C$31"),
        ("=DATE(ANNEE+1,5,15)", "Impôts (IS)",
         "Solde d'IS — mai N+1 (estimation)",
         f"=MAX(0,'{TVA}'!$C$30-4*'{TVA}'!$C$31)"),
        ("=DATE(ANNEE,12,15)", "Impôts (CFE)",
         "CFE annuelle (estimation)", "=CFE_ANNUELLE"),
    ]

    dv_statut = DataValidation(type="list", formula1='"À venir,Payé"',
                               allow_blank=True)
    ws.add_data_validation(dv_statut)
    for i, (d_f, org, nat, amt_f) in enumerate(rows):
        r = first + i
        alt = BEIGE_ALT if i % 2 else BEIGE
        ccalc(ws, f"B{r}", d_f, DATEF, bg=alt).alignment = A_C
        c = ws[f"C{r}"]
        c.value = org
        c.font = F_LABEL
        c.fill = fill(alt)
        c.border = BORDER
        c.alignment = A_L
        c = ws[f"D{r}"]
        c.value = nat
        c.font = F_LABEL
        c.fill = fill(alt)
        c.border = BORDER
        c.alignment = A_L
        ccalc(ws, f"E{r}", amt_f, MONEY, bg=alt)
        cin(ws, f"F{r}", "À venir")
        dv_statut.add(f"F{r}")
        ccalc(
            ws, f"G{r}",
            f'=IF($B{r}="","",IF($F{r}="Payé","✓ Payé",'
            f'IF($B{r}<TODAY(),"Passée",'
            f'IF($B{r}-TODAY()<7,"⚠️ URGENT — dans "&($B{r}-TODAY())&" j",'
            f'"dans "&($B{r}-TODAY())&" j"))))', "@", bg=alt)
        ws[f"G{r}"].alignment = A_C
        # colonne auxiliaire masquée : date si l'échéance est à venir
        # (permet un MIN() universel, compatible partout, sans MINIFS)
        hc = ws[f"H{r}"]
        hc.value = (f'=IF(AND($B{r}<>"",$F{r}<>"Payé",$B{r}>=TODAY()),'
                    f'$B{r},"")')
        hc.font = Font(name=FONT, size=8, color=BEIGE)
        hc.number_format = ";;;"

    ws.conditional_formatting.add(f"G{first}:G{last}", FormulaRule(
        formula=[f'AND($B{first}<>"",$F{first}<>"Payé",'
                 f"$B{first}>=TODAY(),$B{first}-TODAY()<15)"],
        font=Font(name=FONT, size=10, bold=True, color=TERRA)))

    merge(ws, f"B{last + 2}:J{last + 2}",
          "Dates usuelles données à titre indicatif — vérifiez vos dates "
          "exactes sur votre espace impots.gouv.fr et urssaf.fr.",
          F_NOTE, BEIGE, A_LW)
    ws.freeze_panes = "A11"


def build_simulateurs(ws):
    paint(ws, 6, 46)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 48, "C": 18, "D": 20, "E": 3})
    banner(ws, "E", "🧮 SIMULATEURS",
           "En clair : testez vos idées sans risque — salaire, dividendes, "
           "scénarios de fin d'année.")
    tip(ws, "B3:E3",
        "Trois mini-outils indépendants. Modifiez les cellules crème, tout "
        "se recalcule instantanément. Aucune saisie ici ne modifie vos "
        "autres onglets.")

    is_bareme = ("IF({b}<=0,0,IF({b}<=SEUIL_IS,ROUND({b}*TAUX_IS_REDUIT,0),"
                 "ROUND(SEUIL_IS*TAUX_IS_REDUIT+({b}-SEUIL_IS)"
                 "*TAUX_IS_NORMAL,0)))")
    ca_cum = f"SUM('{SUI}'!$D$7:INDEX('{SUI}'!$D$7:$D$18,{MIDX}))"
    dep_cum = f"SUM('{SUI}'!$G$7:INDEX('{SUI}'!$G$7:$G$18,{MIDX}))"

    # --- 1 · combien facturer ? -------------------------------------------
    band(ws, "B5:D5",
         "1 · COMBIEN FACTURER POUR ME PAYER X € NET ? (ESTIMATION)")
    label(ws, "B6", "Salaire net mensuel souhaité")
    cin(ws, "C6", "=SALAIRE_NET", MONEY0)
    label(ws, "B7", "+ Cotisations sociales estimées (selon votre statut)")
    ccalc(ws, "C7", "=ROUND($C$6*TAUX_CHARGES,2)", MONEY0)
    label(ws, "B8", "→ Coût total de votre rémunération", bold=True)
    ccalc(ws, "C8", "=$C$6+$C$7", MONEY0, bold=True)
    label(ws, "B9", "+ Frais de fonctionnement mensuels moyens (auto)")
    ccalc(ws, "C9",
          f"=IFERROR(IF('{DEP}'!$O$6>0,'{DEP}'!$O$6/{MIDX},"
          f"('{BUD}'!$O$31-'{BUD}'!$O$24-'{BUD}'!$O$25)/12),0)", MONEY0)
    label(ws, "B10", "+ Marge de sécurité mensuelle (modifiable)")
    cin(ws, "C10", 300, MONEY0)
    label(ws, "B11", "+ Provision d'IS indicative sur cette marge")
    ccalc(ws, "C11",
          "=IFERROR(IF($C$10*12<=SEUIL_IS,ROUND($C$10*TAUX_IS_REDUIT,2),"
          "ROUND((SEUIL_IS*TAUX_IS_REDUIT+($C$10*12-SEUIL_IS)"
          "*TAUX_IS_NORMAL)/12,2)),0)", MONEY0)
    label(ws, "B12", "→ CA HT à facturer chaque mois", bold=True)
    ws["B12"].font = F_TOTAL
    ccalc(ws, "C12", "=ROUND(SUM($C$8:$C$11),0)", MONEY0, bold=True,
          font=Font(name=FONT, size=13, bold=True, color=BORDEAUX))
    label(ws, "B13", "Jours facturables par mois (modifiable)")
    cin(ws, "C13", 18, "0")
    label(ws, "B14", "→ Tarif jour indicatif (TJM)", bold=True)
    ccalc(ws, "C14", "=IFERROR(ROUND($C$12/$C$13,0),0)", MONEY0,
          bold=True, font=F_TOTAL)
    merge(ws, "B15:D15",
          "Estimation hors TVA. Les frais moyens viennent de vos dépenses "
          "saisies (ou de votre budget en tout début d'année).",
          F_NOTE, BEIGE, A_LW)
    box(ws, 5, 2, 15, 4, TAUPE)

    # --- 2 · rémunération vs dividendes -------------------------------------
    band(ws, "B18:D18",
         "2 · RÉMUNÉRATION OU DIVIDENDES ? (COMPARATEUR INDICATIF)")
    label(ws, "B19",
          "Enveloppe annuelle disponible (avant charges et impôts)")
    cin(ws, "C19", 30000, MONEY0)
    label(ws, "B20", "Option A — Rémunération : net perçu sur l'année",
          bold=True)
    ccalc(ws, "C20", "=ROUND($C$19/(1+TAUX_CHARGES),0)", MONEY0,
          bold=True, font=F_TOTAL)
    merge(ws, "B21:D21",
          "Les cotisations versées financent votre retraite, votre "
          "maladie et votre protection sociale.", F_NOTE, BEIGE, A_LW)
    label(ws, "B22", "Option B — IS payé d'abord par la société")
    ccalc(ws, "C22", "=" + is_bareme.format(b="$C$19"), MONEY0)
    label(ws, "B23", "Option B — Dividendes distribuables (après IS)")
    ccalc(ws, "C23", "=$C$19-$C$22", MONEY0)
    label(ws, "B24", "Option B — Flat tax (PFU, taux en ⚙️ Paramètres)")
    ccalc(ws, "C24", "=ROUND($C$23*FLAT_TAX,0)", MONEY0)
    label(ws, "B25", "Option B — Dividendes : net perçu sur l'année",
          bold=True)
    ccalc(ws, "C25", "=$C$23-$C$24", MONEY0, bold=True, font=F_TOTAL)
    label(ws, "B26", "Verdict indicatif", bold=True)
    merge(ws, "C26:D26",
          '=IF($C$20>=$C$25,"Rémunération : +"&ROUND($C$20-$C$25,0)'
          '&" € net","Dividendes : +"&ROUND($C$25-$C$20,0)&" € net")',
          Font(name=FONT, size=11, bold=True, color=BORDEAUX), BEIGE_ALT,
          A_C, BORDER)
    merge(ws, "B27:D28",
          "Comparateur simplifié à visée pédagogique — l'arbitrage réel "
          "dépend de votre situation (retraite, protection sociale, "
          "abattements, autres revenus…). Parlez-en à un "
          "expert-comptable.", F_NOTE, BEIGE, A_LW)
    box(ws, 18, 2, 28, 4, TAUPE)

    # --- 3 · et si… ? ---------------------------------------------------------
    band(ws, "B30:D30", "3 · ET SI… ? (PROJECTION DE FIN D'ANNÉE)")
    head(ws, "C31", "Votre hypothèse", TAUPE)
    head(ws, "D31", "Au rythme actuel (auto)", TAUPE)
    label(ws, "B32", "CA mensuel pour les mois restants")
    cin(ws, "C32", 8000, MONEY0)
    ccalc(ws, "D32", f"=IFERROR({ca_cum}/{MIDX},0)", MONEY0)
    label(ws, "B33", "Dépenses mensuelles (hors rémunération) restantes")
    cin(ws, "C33", 1500, MONEY0)
    ccalc(ws, "D33", f"=IFERROR('{DEP}'!$O$6/{MIDX},0)", MONEY0)
    label(ws, "B34", "Mois restants dans l'année")
    merge(ws, "C34:D34", f"=12-{MIDX}",
          F_CALC_B, BEIGE_ALT, A_C, BORDER).number_format = "0"
    label(ws, "B35", "CA annuel projeté")
    ccalc(ws, "C35", f"={ca_cum}+$C$32*(12-{MIDX})", MONEY0)
    ccalc(ws, "D35", f"={ca_cum}+$D$32*(12-{MIDX})", MONEY0)
    label(ws, "B36", "Dépenses annuelles projetées (avec rémunération)")
    ccalc(ws, "C36",
          f"={dep_cum}+($C$33+ROUND(SALAIRE_NET*(1+TAUX_CHARGES),2))"
          f"*(12-{MIDX})", MONEY0)
    ccalc(ws, "D36",
          f"={dep_cum}+($D$33+ROUND(SALAIRE_NET*(1+TAUX_CHARGES),2))"
          f"*(12-{MIDX})", MONEY0)
    label(ws, "B37", "Résultat annuel projeté", bold=True)
    ccalc(ws, "C37", "=$C$35-$C$36", MONEY0, bold=True, font=F_TOTAL)
    ccalc(ws, "D37", "=$D$35-$D$36", MONEY0, bold=True, font=F_TOTAL)
    label(ws, "B38", "IS estimé sur ce résultat")
    ccalc(ws, "C38", "=" + is_bareme.format(b="$C$37"), MONEY0)
    ccalc(ws, "D38", "=" + is_bareme.format(b="$D$37"), MONEY0)
    label(ws, "B39", "Trésorerie estimée fin d'année (hors TVA)",
          bold=True)
    ccalc(ws, "C39",
          f"=IFERROR(INDEX('{TRE}'!$C$20:$N$20,{MIDX})"
          f"+($C$35-{ca_cum})-($C$36-{dep_cum})-$C$38,0)", MONEY0,
          bold=True, font=F_TOTAL)
    ccalc(ws, "D39",
          f"=IFERROR(INDEX('{TRE}'!$C$20:$N$20,{MIDX})"
          f"+($D$35-{ca_cum})-($D$36-{dep_cum})-$D$38,0)", MONEY0,
          bold=True, font=F_TOTAL)
    ws.conditional_formatting.add("C39:D39", CellIsRule(
        operator="lessThan", formula=["0"],
        font=Font(name=FONT, size=10, bold=True, color=TERRA)))
    merge(ws, "B40:D41",
          "Projection simplifiée : elle suppose des encaissements "
          "immédiats et ne tient pas compte de la TVA. Utilisez-la pour "
          "comparer des scénarios, pas comme une prévision exacte.",
          F_NOTE, BEIGE, A_LW)
    box(ws, 30, 2, 41, 4, TAUPE)


LEXIQUE = [
    ("HT / TTC", "HT = hors taxes, TTC = toutes taxes comprises. Ex. : "
     "1 000 € HT à 20 % de TVA = 1 200 € TTC."),
    ("TVA collectée", "La TVA que vous facturez à vos clients pour le "
     "compte de l'État. Ex. : sur une facture de 1 200 € TTC à 20 %, la "
     "TVA collectée est de 200 €."),
    ("TVA déductible", "La TVA payée sur vos achats professionnels, que "
     "vous pouvez récupérer. Ex. : 24 € de TVA sur un logiciel à "
     "120 € HT."),
    ("TVA à décaisser", "TVA collectée − TVA déductible = ce que vous "
     "reversez réellement à l'État."),
    ("Franchise en base de TVA", "Régime où vous ne facturez pas de TVA "
     "(sous certains seuils de CA). Plus simple, mais TVA non récupérable "
     "sur vos achats."),
    ("Réel simplifié", "Régime de TVA avec 2 acomptes par an (juillet et "
     "décembre) et une déclaration annuelle (CA12)."),
    ("Réel normal", "Régime de TVA avec déclaration et paiement chaque "
     "mois (ou chaque trimestre)."),
    ("IS (impôt sur les sociétés)", "L'impôt payé par la société sur ses "
     "bénéfices : 15 % jusqu'à 42 500 € sous conditions, 25 % au-delà "
     "(taux indicatifs, modifiables en ⚙️ Paramètres)."),
    ("Acompte", "Un paiement anticipé. L'IS se paie en 4 acomptes "
     "trimestriels, régularisés l'année suivante."),
    ("CFE", "Cotisation foncière des entreprises : taxe locale annuelle "
     "due par presque toutes les sociétés, payée au 15 décembre."),
    ("URSSAF", "L'organisme qui collecte les cotisations sociales "
     "(retraite, maladie, famille…)."),
    ("Cotisations sociales", "Les prélèvements calculés sur la "
     "rémunération, qui financent votre protection sociale."),
    ("Assimilé salarié", "Statut du président de SASU : protection "
     "sociale proche de celle d'un salarié, cotisations plus élevées "
     "(≈ 82 % du net, indicatif)."),
    ("TNS", "Travailleur non salarié : statut du gérant majoritaire de "
     "SARL. Cotisations plus faibles (≈ 45 % du net, indicatif), "
     "protection différente."),
    ("Charge déductible", "Une dépense professionnelle qui réduit votre "
     "bénéfice imposable. Ex. : logiciel, loyer, honoraires."),
    ("Amortissement", "Étaler le coût d'un gros achat (ordinateur, "
     "véhicule…) sur plusieurs années comptables."),
    ("Immobilisation", "Un bien durable acheté par l'entreprise, qui "
     "s'amortit sur plusieurs années au lieu de passer en charge d'un "
     "coup."),
    ("Trésorerie", "L'argent réellement disponible sur le compte de la "
     "société. On peut être rentable et à sec : surveillez-la."),
    ("Runway", "Le nombre de mois pendant lesquels vous pouvez tenir "
     "avec la trésorerie actuelle, au rythme de dépenses actuel."),
    ("BFR", "Besoin en fonds de roulement : l'argent immobilisé entre le "
     "moment où vous payez vos charges et celui où vos clients vous "
     "paient."),
    ("Marge nette", "Résultat ÷ chiffre d'affaires. Ex. : 12 000 € de "
     "résultat pour 100 000 € de CA = 12 % de marge nette."),
    ("Seuil de rentabilité", "Le chiffre d'affaires minimum pour couvrir "
     "toutes vos charges. En dessous, vous perdez de l'argent."),
    ("Prévisionnel", "Ce que vous prévoyez (votre budget). Le "
     "« réalisé », c'est ce qui s'est vraiment passé."),
    ("Écart", "Réalisé − prévu. Dans ce fichier, un écart positif est "
     "toujours favorable."),
    ("Provision", "De l'argent mis de côté pour une dépense certaine à "
     "venir (TVA, URSSAF, IS…). La clé pour dormir tranquille."),
    ("Dividende", "La part du bénéfice (après IS) versée aux associés, "
     "soumise à la flat tax."),
    ("Flat tax (PFU)", "Prélèvement forfaitaire unique de 30 % "
     "(indicatif) sur les dividendes : impôt + prélèvements sociaux en "
     "une fois."),
    ("Exercice", "La période comptable de 12 mois de la société, souvent "
     "l'année civile."),
    ("Note de frais", "Le document qui permet de vous faire rembourser "
     "une dépense professionnelle payée avec votre argent personnel."),
    ("Indemnité kilométrique", "Remboursement forfaitaire de vos trajets "
     "professionnels effectués avec votre véhicule personnel, selon le "
     "barème officiel."),
    ("Comptabilité d'engagement", "On enregistre les factures à leur "
     "date d'émission, pas à leur date de paiement (contraire : "
     "comptabilité de trésorerie)."),
    ("Rapprochement bancaire", "Vérifier que votre suivi correspond "
     "bien, ligne à ligne, au relevé de la banque."),
]


def build_lexique(ws):
    paint(ws, 4, len(LEXIQUE) + 12)
    ws.sheet_properties.tabColor = TAUPE
    widths(ws, {"A": 2.5, "B": 30, "C": 100, "D": 3})
    banner(ws, "C", "📖 LEXIQUE",
           "En clair : tous les mots de la gestion expliqués simplement, "
           "avec des exemples.")
    tip(ws, "B3:C3",
        "Un terme vous échappe dans le fichier ou chez votre banquier ? "
        "Il est expliqué ici, sans jargon.")
    head(ws, "B5", "Terme", TAUPE)
    head(ws, "C5", "En clair", TAUPE)
    for i, (terme, defi) in enumerate(LEXIQUE):
        r = 6 + i
        alt = BEIGE_ALT if i % 2 else BEIGE
        c = ws[f"B{r}"]
        c.value = terme
        c.font = F_LABEL_B
        c.fill = fill(alt)
        c.border = BORDER
        c.alignment = Alignment(horizontal="left", vertical="top")
        c = ws[f"C{r}"]
        c.value = defi
        c.font = F_LABEL
        c.fill = fill(alt)
        c.border = BORDER
        c.alignment = A_LW
        ws.row_dimensions[r].height = 28
    ws.freeze_panes = "A6"


# ===================================================================== main ==

def main():
    wb = Workbook()
    ws_dash = wb.active
    ws_dash.title = DASH
    for name in (START, PAR, BUD, REC, DEP, SUI, TRE, TVA, ECH, SIM, KM,
                 LEX):
        wb.create_sheet(name)

    build_parametres(wb[PAR])
    define_names(wb)
    build_start(wb[START])
    build_budget(wb[BUD])
    build_recettes(wb[REC])
    build_depenses(wb[DEP])
    build_suivi(wb[SUI])
    build_tresorerie(wb[TRE])
    build_tva(wb[TVA])
    build_echeancier(wb[ECH])
    build_simulateurs(wb[SIM])
    build_km(wb[KM])
    build_lexique(wb[LEX])
    build_dashboard(ws_dash)

    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    props = wb.properties
    props.title = "Gestion SASU/SARL 2026"
    props.creator = "La Fabrique Astucieuse"
    props.description = ("Tableau de bord financier complet pour SASU et "
                         "SARL : budget, recettes, dépenses, trésorerie, "
                         "TVA, IS, URSSAF, CFE et indemnités kilométriques.")

    out = "Gestion-SASU-SARL-2026.xlsx"
    wb.save(out)
    print(f"OK — {out} généré.")


if __name__ == "__main__":
    main()
