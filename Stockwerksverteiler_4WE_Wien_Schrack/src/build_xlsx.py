# -*- coding: utf-8 -*-
"""Erzeugt Auslegung + Stückliste + Kabelliste (xlsx) aus data.py.  Aufruf: python3 build_xlsx.py <out.xlsx>"""
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import data as D

F = "Arial"
F_N = Font(name=F, size=9)
F_B = Font(name=F, size=9, bold=True)
F_H = Font(name=F, size=12, bold=True)
F_IN = Font(name=F, size=9, color="0000FF")
F_LINK = Font(name=F, size=9, color="008000")
F_G = Font(name=F, size=8, color="666666")
YEL = PatternFill("solid", fgColor="FFFF00")
HDR = PatternFill("solid", fgColor="D9E1F2")
SEC = PatternFill("solid", fgColor="EDEDED")
OPT = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WR = Alignment(wrap_text=True, vertical="top")
CEN = Alignment(horizontal="center", vertical="top")

DS_URL = "https://www.schrack.com/shop/catalog/download/deliverOnlineDataSheet?articlenr={}&realUser=yes"
SRC_K = {
    "IS504037--": "Zubehörliste TYTAN T (schrack.com, IS503031)",
    "IS504036--": "Zubehörliste TYTAN T (schrack.com, IS503031)",
    "IS504136--": "Zubehörliste TYTAN T (schrack.com, IS503031)",
    "IS504135--": "Zubehörliste TYTAN T (schrack.com, IS503031)",
    "IKB04035--": "schrack.com Shop: main-branch-terminal-4-pole-125a-in-1x35mm2-out-10x16mm2-ikb04035",
    "SI022270--": "Schrack-Katalog Modul 2000 (k_il160121vs_de.pdf), Einbauteile-Konfiguration",
}


def cell(ws, ref, val, font=F_N, fill=None, fmt=None, al=None, border=True):
    c = ws[ref]
    c.value = val
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if al:
        c.alignment = al
    if border:
        c.border = BOX
    return c


# ============================================================================ Info
def sheet_info(wb):
    ws = wb.active
    ws.title = "Info"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 110
    ws["A1"] = "Stückliste und Auslegung – Stockwerksverteiler + Wohnungsverteiler Top 1–4, Wien"
    ws["A1"].font = F_H
    rows = [
        ("Projekt", D.PROJECT),
        ("Planung", f"{D.COMPANY} · Plan-Nr. {D.PLAN_NO} · Stand {D.DATE}"),
        ("Zugehöriger Plan", "Stromlaufplan_SV-WV_Wien_4WE.pdf (33 Blätter, A3, allpolig)"),
        ("Netzbetreiber", D.NETZ + " – Smart Meter Beistellung, Plombierung, Zählersteckklemmen nach OVE R 21:2019"),
        ("Hersteller", "Schrack Technik (Eigenmarke) und über Schrack vertriebene Unterlieferanten (Phoenix Contact)"),
        ("", ""),
        ("Legende", ""),
        ("gelbe Zelle, blaue Schrift", "Eingabe / Annahme – darf geändert werden (Preise, Rabatte, Optionen, Längen, Lastannahmen)"),
        ("schwarze Schrift", "Formel – nicht überschreiben"),
        ("grüne Schrift", "Verweis auf ein anderes Blatt"),
        ("Status verifiziert", "Artikelnummer am 25.11.2026 auf schrack.com (Datenblatt bzw. Shopseite) geprüft"),
        ("Status Schrack-Katalog", "Nummer aus Schrack-Zubehörliste/-Konfiguration; kein Online-Datenblatt abrufbar → vor Bestellung bestätigen"),
        ("Status Option", "Spar-/Zusatzvariante, über Spalte 'Option aktiv' (1/0) zu- oder abschaltbar"),
        ("Status anfragen", "Meterware/Kleinmaterial: Schrack-Art.-Nr. lt. Kabel-/Installationssortiment anfragen"),
        ("Beispiel Preiseingabe", "Listenpreis 12,50 €/Stk · Rabatt 35 % → Netto 8,13 €/Stk (Format: Zahl ohne €-Zeichen, Rabatt in %)"),
        ("", ""),
        ("Wichtige Annahmen", ""),
        ("Lastannahme", "14,5 kW je Wohnung (ohne el. Warmwasser), Gleichzeitigkeit 0,60 für 4 WE, Allgemein 2 kW, Reserve 11 kW"),
        ("Leitungslängen", "W0 30 m, W1–W4 ≤ 25 m, Endstromkreise lt. Kabelliste – durch Installationsplan ersetzen"),
        ("Kurzschluss", "Ik3 am Stockwerksverteiler ≤ 10 kA (bei Wiener Netze anfragen); Zs am WV 0,35 Ω (Annahme)"),
        ("Preise", "keine Listenpreise hinterlegt (Eingabefelder); Schrack-Konditionen eintragen"),
        ("Normen", "OVE E 8101, TAEV 2020 + Ausführungsbest. Wiener Netze, ÖVE/ÖNORM EN 61439-1/-3, OVE R 21:2019"),
    ]
    for i, (a, b) in enumerate(rows, start=3):
        ws.cell(i, 1, a).font = F_B if b == "" or a.startswith(("gelbe", "Legende", "Wichtige")) else F_B
        ws.cell(i, 2, b).font = F_N
        ws.cell(i, 2).alignment = WR
        if a == "gelbe Zelle, blaue Schrift":
            ws.cell(i, 1).fill = YEL
            ws.cell(i, 1).font = F_IN
        if a in ("Legende", "Wichtige Annahmen"):
            ws.cell(i, 1).fill = SEC
            ws.cell(i, 2).fill = SEC


# ============================================================================ Auslegung
INP = {}


def sheet_ausl(wb):
    ws = wb.create_sheet("Auslegung")
    for col, w in zip("ABCDE", (46, 14, 10, 62, 14)):
        ws.column_dimensions[col].width = w
    ws["A1"] = "Auslegung Steigleitung, Zuleitungen, Endstromkreise, Verlustleistung"
    ws["A1"].font = F_H
    r = 3
    for c, h in zip("ABCD", ("Eingabe / Annahme", "Wert", "Einheit", "Quelle / Bemerkung")):
        cell(ws, f"{c}{r}", h, F_B, HDR)
    i = D.INPUT
    inputs = [
        ("P_WE", "Leistungsbedarf je Wohnung (ohne el. WW)", i["P_WE"], "kW", "Annahme, Anhalt DIN 18015-1"),
        ("g_WE", "Gleichzeitigkeitsfaktor 4 Wohnungen", i["g_WE"], "–", "Annahme"),
        ("n_WE", "Anzahl Wohnungen", 4, "–", "Vorgabe"),
        ("P_ALLG", "Allgemein (Stiegenhaus, Sprechanlage, RWA)", i["P_ALLG"], "kW", "Annahme"),
        ("P_RES", "Reserve-Zählerplatz (z. B. Wallbox 11 kW)", i["P_RES"], "kW", "Annahme"),
        ("g_RES", "Gleichzeitigkeit Reserve", i["g_RES"], "–", "Annahme"),
        ("COS", "Leistungsfaktor cos φ", D.COSPHI, "–", "Annahme"),
        ("UN", "Nennspannung Un (verkettet)", D.U_N, "V", "Netz"),
        ("U0", "Spannung gegen Erde U0", D.U_0, "V", "Netz"),
        ("RHO", "spez. Widerstand Cu (Betrieb) ρ", D.CU_RHO, "Ω·mm²/m", "1,25·ρ20 (IEC 60364-5-52 Anh. G)"),
        ("XL", "Reaktanzbelag x'", D.X_PER_M, "Ω/m", "0,08 mΩ/m"),
        ("L_W0", "Länge -W0 HV → SV", i["L_W0"], "m", "Annahme"),
        ("S_W0", "Querschnitt -W0", 25, "mm²", "Vorgabe Auftraggeber (5×25)"),
        ("IN_HV", "Vorsicherung NH00 gG im HV", 80, "A", "Vorgabe Auftraggeber"),
        ("IZ_W0", "Iz -W0 (PVC 70 °C, Verlegeart C, 3 belastete Adern)", 96, "A", "OVE E 8101 / IEC 60364-5-52 Tab. B.52.4"),
        ("L_WT", "Länge -W1…-W4 SV → WV (Höchstwert)", i["L_WTOP"], "m", "Vorgabe ≤ 25 m"),
        ("S_WT", "Querschnitt -W1…-W4", 10, "mm²", "Vorgabe Auftraggeber (5×10)"),
        ("IN_VZ", "Vorzählersicherung NEOZED D02 gG", 35, "A", "Vorgabe Auftraggeber"),
        ("IZ_WT", "Iz -W1…-W4 (Verlegeart B2, 3 belastete Adern)", 46, "A", "IEC 60364-5-52 Tab. B.52.4"),
        ("IK_SV", "Ik3 max. am Stockwerksverteiler", i["IK_SV"], "kA", "Annahme – bei Wiener Netze anfragen"),
        ("Z_WV", "Schleifenimpedanz Zs am Wohnungsverteiler", i["Z_WV"], "Ω", "Annahme – bei IBN messen"),
        ("RDF_SV", "Bemessungsbelastungsfaktor RDF SV (6–9 Abgänge)", i["RDF_SV"], "–", "EN 61439-3 Tab. 101"),
        ("RDF_WV", "Bemessungsbelastungsfaktor RDF WV (≥ 10 Abgänge)", i["RDF_WV"], "–", "EN 61439-3 Tab. 101"),
        ("PV_WIRE", "Verlustleistung Verdrahtung SV (Zuschlag)", 10, "W", "Richtwert"),
        ("PVMAX_SV", "zul. Verlustleistung Gehäuse SV IL166333WS", 455, "W", "Schrack-Datenblatt"),
        ("PVMAX_WV", "zul. Verlustleistung Gehäuse WV BK0851042F", 54, "W", "Schrack-Datenblatt"),
        ("DU_BUD", "ΔU-Budget Endstromkreis", 0.03, "%", "Planungswert (Summe ab Zähler ≤ 4 %)"),
        ("L_BUS", "genutzte Länge Phasenschiene je WV", 0.40, "m", "21 TE · 17,8 mm + Verschnitt"),
    ]
    r = 4
    for key, txt, val, unit, src in inputs:
        cell(ws, f"A{r}", txt)
        fmt = "0.0%" if unit == "%" else ("0.00000" if key in ("RHO", "XL") else "0.00")
        cell(ws, f"B{r}", val, F_IN, YEL, fmt)
        cell(ws, f"C{r}", "–" if unit == "%" else unit)
        cell(ws, f"D{r}", src, F_G)
        INP[key] = f"Auslegung!$B${r}"
        INP["_" + key] = f"$B${r}"
        r += 1
    r += 1
    for c, h in zip("ABCD", ("Berechnung", "Ergebnis", "Einheit", "Formel / Prüfung")):
        cell(ws, f"{c}{r}", h, F_B, HDR)
    r += 1
    L = lambda k: INP["_" + k]
    calc = [
        ("P", "Bemessungsleistung P", f"={L('n_WE')}*{L('P_WE')}*{L('g_WE')}+{L('P_ALLG')}+{L('P_RES')}*{L('g_RES')}",
         "kW", "n·P_WE·g + P_Allg + P_Res·g_Res", "0.0"),
        ("IB", "Betriebsstrom Ib", "=B{P}*1000/(SQRT(3)*{UN}*{COS})", "A", "P/(√3·Un·cos φ)", "0.0"),
        ("CHK1", "Prüfung Ib ≤ In (NH00)", "=IF(B{IB}<={IN_HV},\"OK\",\"PRÜFEN\")", "", "Ib ≤ In", None),
        ("CHK2", "Prüfung Überlast -W0: 1,6·In ≤ 1,45·Iz", "=IF(1.6*{IN_HV}<=1.45*{IZ_W0},\"OK\",\"PRÜFEN – 5×35 mm²\")", "",
         "gG: I2 = 1,6·In", None),
        ("DUW0V", "Spannungsfall -W0", "=SQRT(3)*B{IB}*({RHO}*{L_W0}/{S_W0}*{COS}+{XL}*{L_W0}*SQRT(1-{COS}^2))", "V",
         "√3·Ib·(R·cos φ + X·sin φ)", "0.00"),
        ("DUW0", "Spannungsfall -W0 relativ", "=B{DUW0V}/{UN}", "%", "ΔU/Un", "0.00%"),
        ("LW005", "max. Länge -W0 für ΔU 0,5 %", "=0.005/(B{DUW0}/{L_W0})", "m", "bei Ib, linear", "0.0"),
        ("DUWTV", "Spannungsfall -W1…-W4 (bei In = 35 A)", "=SQRT(3)*{IN_VZ}*({RHO}*{L_WT}/{S_WT}*{COS}+{XL}*{L_WT}*SQRT(1-{COS}^2))",
         "V", "√3·In·(R·cos φ + X·sin φ)", "0.00"),
        ("DUWT", "Spannungsfall -W1…-W4 relativ", "=B{DUWTV}/{UN}", "%", "ΔU/Un", "0.00%"),
        ("CHK3", "Prüfung Überlast -W1…-W4: 1,6·In ≤ 1,45·Iz", "=IF(1.6*{IN_VZ}<=1.45*{IZ_WT},\"OK\",\"PRÜFEN\")", "",
         "gG: I2 = 1,6·In", None),
        ("SEL", "Selektivität NH00 → D02 (Verhältnis)", "={IN_HV}/{IN_VZ}", "–", "≥ 1,6 : 1 selektiv", "0.00"),
        ("CHK4", "Prüfung Selektivität", "=IF(B{SEL}>=1.6,\"OK\",\"PRÜFEN\")", "", "", None),
        ("IKWV", "Ik3 am Wohnungsverteiler", "={U0}/SQRT(({U0}/({IK_SV}*1000)+0.01786*{L_WT}/{S_WT})^2+({XL}*{L_WT})^2)/1000",
         "kA", "über -W1 (20 °C)", "0.00"),
        ("CHK5", "Prüfung Schaltvermögen WV (FI 6 kA)", "=IF(B{IKWV}<=6,\"OK\",\"PRÜFEN\")", "", "", None),
        ("PVSV", "Verlustleistung SV bei Nennstrom", "=SUMPRODUCT((Stückliste!$B$9:$B$200=\"SV\")*Stückliste!$F$9:$F$200*"
                                                   "Stückliste!$L$9:$L$200*Stückliste!$M$9:$M$200*Stückliste!$Q$9:$Q$200)+{PV_WIRE}",
         "W", "Σ Menge·Pv·Faktor + Verdrahtung", "0.0"),
        ("PVSVB", "Verlustleistung SV im Betrieb", "=B{PVSV}*{RDF_SV}^2", "W", "Pv·RDF²", "0.0"),
        ("CHK6", "Prüfung Erwärmung SV", "=IF(B{PVSVB}<={PVMAX_SV},\"OK\",\"PRÜFEN\")", "", "≤ Pv,max", None),
        ("PVWV", "Verlustleistung je WV bei Nennstrom", "=SUMPRODUCT((Stückliste!$B$9:$B$200=\"WV1–4\")*Stückliste!$D$9:$D$200*"
                                                      "Stückliste!$L$9:$L$200*Stückliste!$M$9:$M$200*Stückliste!$Q$9:$Q$200)",
         "W", "Σ Menge je WV·Pv·Faktor", "0.0"),
        ("PVWVB", "Verlustleistung je WV im Betrieb", "=B{PVWV}*{RDF_WV}^2", "W", "Pv·RDF²", "0.0"),
        ("CHK7", "Prüfung Erwärmung WV", "=IF(B{PVWVB}<={PVMAX_WV},\"OK\",\"PRÜFEN\")", "", "≤ Pv,max", None),
    ]
    rowof = {}
    for key, *_ in calc:
        rowof[key] = r
        r += 1
    r = rowof[calc[0][0]]
    fmtmap = {k: INP["_" + k] for k in [x[0] for x in inputs]}
    for key, txt, f, unit, rem, fmt in calc:
        ff = f
        for k2, rr in rowof.items():
            ff = ff.replace("{" + k2 + "}", str(rr))
        for k2, ref in fmtmap.items():
            ff = ff.replace("{" + k2 + "}", ref)
        cell(ws, f"A{r}", txt)
        c = cell(ws, f"B{r}", ff, F_B if key.startswith("CHK") else F_N, fmt=fmt)
        cell(ws, f"C{r}", "–" if unit == "%" else unit)
        cell(ws, f"D{r}", rem, F_G)
        INP["R_" + key] = r
        r += 1
    # ---- Endstromkreise
    r += 1
    for c, h in zip("ABCDE", ("Endstromkreis", "Zs,max 0,4 s [Ω]", "In [A]", "Lmax Abschaltung [m] · Lmax ΔU-Budget [m]",
                              "S [mm²]")):
        cell(ws, f"{c}{r}", h, F_B, HDR)
    r += 1
    for name, In, S, ph in (("B10 · NYM-J 3×1,5 (1-phasig)", 10, 1.5, 1), ("B16 · NYM-J 3×2,5 (1-phasig)", 16, 2.5, 1),
                            ("B16 3-pol. · NYM-J 5×2,5 (Kochfeld)", 16, 2.5, 3)):
        cell(ws, f"A{r}", name)
        cell(ws, f"C{r}", In, F_IN, YEL)
        cell(ws, f"E{r}", S, F_IN, YEL)
        cell(ws, f"B{r}", f"={L('U0')}/(5*C{r})", fmt="0.00")
        if ph == 1:
            ldu = f"{L('DU_BUD')}*{L('U0')}/(2*C{r}*{L('RHO')}/E{r})"
        else:
            ldu = f"{L('DU_BUD')}*{L('UN')}/(SQRT(3)*C{r}*{L('RHO')}/E{r})"
        cell(ws, f"D{r}", f"=TEXT((B{r}-{L('Z_WV')})/(2*{L('RHO')}/E{r}),\"0\")&\" m  ·  \"&TEXT({ldu},\"0\")&\" m\"")
        r += 1
    cell(ws, f"A{r}", "Maßgebend ist der Spannungsfall (Budget lt. Eingabe); Abschaltbedingung (Nullung) bei IBN messen.",
         F_G, border=False)
    # ---- Sensitivität
    r += 2
    cell(ws, f"A{r}", "Sensitivität Ib [A] – Gleichzeitigkeit (Zeilen) × Reserveleistung (Spalten)", F_B, SEC)
    r += 1
    gs = [0.5, 0.6, 0.7, 0.8, 1.0]
    ps = [0, 11, 22]
    cell(ws, f"A{r}", "g_WE  \\  P_Res [kW]", F_B, HDR)
    for j, pv in enumerate(ps):
        cell(ws, f"{get_column_letter(2 + j)}{r}", pv, F_IN, YEL, "0")
    hr = r
    r += 1
    for g in gs:
        cell(ws, f"A{r}", g, F_IN, YEL, "0.00")
        for j in range(len(ps)):
            col = get_column_letter(2 + j)
            cell(ws, f"{col}{r}", f"=({L('n_WE')}*{L('P_WE')}*$A{r}+{L('P_ALLG')}+{col}${hr})*1000/(SQRT(3)*{L('UN')}*{L('COS')})",
                 fmt="0.0")
        r += 1
    cell(ws, f"A{r}", "Werte > In NH00 (80 A) → Vorsicherung/Querschnitt erhöhen.", F_G, border=False)
    r += 2
    cell(ws, f"A{r}", "Sensitivität ΔU -W0 [%] bei Ib – Länge (Zeilen) × Querschnitt (Spalten)", F_B, SEC)
    r += 1
    cell(ws, f"A{r}", "L [m]  \\  S [mm²]", F_B, HDR)
    ss = [25, 35, 50]
    for j, sv in enumerate(ss):
        cell(ws, f"{get_column_letter(2 + j)}{r}", sv, F_IN, YEL, "0")
    hr = r
    r += 1
    for Lm in (15, 20, 25, 30, 35, 40):
        cell(ws, f"A{r}", Lm, F_IN, YEL, "0")
        for j in range(len(ss)):
            col = get_column_letter(2 + j)
            cell(ws, f"{col}{r}", f"=SQRT(3)*$B${INP['R_IB']}*({L('RHO')}*$A{r}/{col}${hr}*{L('COS')}+{L('XL')}*$A{r}*"
                                  f"SQRT(1-{L('COS')}^2))/{L('UN')}", fmt="0.00%")
        r += 1
    ws.freeze_panes = "A4"


# ============================================================================ Stückliste
def sheet_bom(wb):
    ws = wb.create_sheet("Stückliste")
    heads = ["Pos", "Anlage", "BMK", "Menge je Anlage", "Anzahl Anlagen", "Menge gesamt", "Einheit", "Bezeichnung",
             "Schrack-Art.-Nr.", "Hersteller", "Status", "Pv je Einheit [W]", "Pv-Faktor", "Listenpreis €/Einh.",
             "Rabatt %", "Netto €/Einh.", "Option aktiv (1/0)", "Summe netto €", "Bemerkung", "Quelle"]
    widths = [5, 9, 24, 9, 8, 9, 6, 62, 14, 22, 18, 9, 8, 11, 8, 11, 9, 12, 34, 40]
    for k, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(k)].width = w
    ws["A1"] = "Stückliste – Schrack Technik und Unterlieferanten (Stockwerksverteiler + Wohnungsverteiler Top 1–4)"
    ws["A1"].font = F_H
    ws["A2"] = ("Gelbe Felder ausfüllen: Listenpreis, Rabatt, Option aktiv, Mengen Meterware. Beispiel: Listenpreis 12,50 · "
                "Rabatt 35 % → Netto 8,13 €. Status 'Schrack-Katalog' = Nummer vor Bestellung bestätigen.")
    ws["A2"].font = F_G
    ws["A3"] = "Summe SV netto €"
    ws["A4"] = "Summe WV (×4) netto €"
    ws["A5"] = "Gesamtsumme netto €"
    for rr in (3, 4, 5):
        ws.merge_cells(f"A{rr}:C{rr}")
        ws[f"A{rr}"].font = F_B
    ws["D3"] = '=SUMIFS($R$9:$R$200,$B$9:$B$200,"SV")'
    ws["D4"] = '=SUMIFS($R$9:$R$200,$B$9:$B$200,"WV1–4")'
    ws["D5"] = "=D3+D4"
    for rr in (3, 4, 5):
        ws[f"D{rr}"].number_format = '#,##0.00 "€"'
        ws[f"D{rr}"].font = F_B
    hr = 8
    for k, h in enumerate(heads, start=1):
        c = ws.cell(hr, k, h)
        c.font = F_B
        c.fill = HDR
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[hr].height = 30
    sv, wv = D.bom()
    opt_rows = [
        ("SV", "Option -6F1", 3, "Stk", "IS504037--", "Sicherungseinsätze für Reserve-Zählerplatz (erst bei Belegung)", 0),
        ("SV", "Option -W0", D.INPUT["L_W0"], "m", "NHXHX-J 5x35",
         "Alternative bei Verlegeart B2/Häufung: 5×35 statt 5×25 – dann Zeile -W0 (5×25) auf 0 setzen", 0),
    ]
    D.ART.setdefault("NHXHX-J 5x35", dict(t="Kabel NHXHX-J 5×35 mm² RM, halogenfrei", m="lt. Schrack-Kabelsortiment",
                                          pv=0.0, st="A"))
    allrows = [(a, b, q, u, art, rem, None) for a, b, q, u, art, rem in sv]
    allrows += [(a, b, q, u, art, rem, f) for a, b, q, u, art, rem, f in opt_rows]
    allrows += [(a, b, q, u, art, rem, None) for a, b, q, u, art, rem in wv]
    dv = DataValidation(type="list", formula1='"0,1"', allow_blank=False)
    ws.add_data_validation(dv)
    r = hr + 1
    stxt = {"V": "verifiziert", "K": "Schrack-Katalog", "O": "Option", "A": "Art.-Nr. anfragen", "B": "Beistellung WN"}
    for pos, (a, b, q, u, art, rem, flag) in enumerate(allrows, start=1):
        A = D.ART[art]
        isopt = A["st"] == "O" or flag is not None
        anl = "SV" if a == "SV" else "WV1–4"
        vals = {
            1: pos, 2: anl, 3: b, 4: q, 5: 1 if a == "SV" else 4, 6: f"=D{r}*E{r}", 7: u, 8: A["t"],
            9: art if A["st"] not in ("A", "B") else "—", 10: A["m"], 11: "Option" if isopt else stxt[A["st"]],
            12: A["pv"], 13: (D.INPUT["L_BUS"] if art == "BSB90116ME" else (0 if A["st"] in ("A",) else 1)),
            14: None, 15: None, 16: f"=IF(N{r}=\"\",\"\",N{r}*(1-O{r}))", 17: (0 if isopt else 1),
            18: f"=IF(P{r}=\"\",0,F{r}*P{r}*Q{r})", 19: rem + ((" · " + A["note"]) if A.get("note") else ""),
            20: (DS_URL.format(art) if A["st"] == "V" and not art.startswith(("NH", "NY", "H07")) else SRC_K.get(art, "")),
        }
        for k, v in vals.items():
            c = ws.cell(r, k, v)
            c.font = F_N
            c.border = BOX
            c.alignment = WR if k in (3, 8, 19, 20) else Alignment(vertical="top")
        for k in (4, 14, 15, 17):
            ws.cell(r, k).fill = YEL
            ws.cell(r, k).font = F_IN
        if art == "BSB90116ME":
            ws.cell(r, 13).fill = YEL
            ws.cell(r, 13).font = F_IN
            ws.cell(r, 13).comment = Comment("Pv-Faktor = genutzte Länge in m (Pv 21,5 W/m bei Nennstrom)", "TBH")
        ws.cell(r, 4).number_format = "0.##"
        ws.cell(r, 6).number_format = "0.##"
        ws.cell(r, 12).number_format = "0.00"
        ws.cell(r, 13).number_format = "0.00"
        ws.cell(r, 14).number_format = "#,##0.00"
        ws.cell(r, 15).number_format = "0%"
        ws.cell(r, 16).number_format = "#,##0.00"
        ws.cell(r, 18).number_format = "#,##0.00"
        dv.add(ws.cell(r, 17))
        if isopt:
            for k in (1, 2, 3, 7, 8, 9, 10, 11):
                ws.cell(r, k).fill = OPT
        r += 1
    last = r - 1
    ws.cell(r + 1, 8, "Summe netto € (aktive Positionen)").font = F_B
    ws.cell(r + 1, 18, f"=SUM(R{hr + 1}:R{last})").font = F_B
    ws.cell(r + 1, 18).number_format = '#,##0.00 "€"'
    ws.freeze_panes = ws.cell(hr + 1, 4)
    ws.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{last}"
    return last


# ============================================================================ Kabelliste
def sheet_cables(wb):
    ws = wb.create_sheet("Kabelliste")
    heads = ["Anlage", "Kabel", "von", "nach", "Typ", "Adern × mm²", "Querschnitt mm²", "Länge m", "I für ΔU [A]",
             "Phasen", "ΔU [V]", "ΔU [%]", "ΔU ab Zähler [%]", "Funktion"]
    widths = [8, 9, 18, 18, 10, 11, 10, 9, 10, 7, 9, 9, 12, 60]
    for k, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(k)].width = w
    ws["A1"] = "Kabelliste mit Spannungsfall je Leitung (Längen = Annahmen, gelb)"
    ws["A1"].font = F_H
    hr = 3
    for k, h in enumerate(heads, start=1):
        c = ws.cell(hr, k, h)
        c.font = F_B
        c.fill = HDR
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[hr].height = 28
    curr = {c["bmk"].replace("-F", "-W"): c["In"] for c in D.WV_G1 + D.WV_G2 + [D.WV_F3]}
    allg_i = {c["cable"]: c["In"] for c in D.ALLG if c["cable"]}
    r = hr + 1
    rho, cos, xl, un, u0 = INP["RHO"], INP["COS"], INP["XL"], INP["UN"], INP["U0"]
    for a, n, f, t, typ, cs, L, fn in D.cables():
        s = None
        ph = 1
        I = None
        if cs not in ("—",):
            nc, sq = cs.split("x")
            s = float(sq.replace(",", "."))
        if n == "-W0":
            I, ph = f"=Auslegung!$B${INP['R_IB']}", 3
        elif a == "SV" and n in ("-W1", "-W2", "-W3", "-W4"):
            I, ph = f"={INP['IN_VZ']}", 3
        elif a == "SV":
            I = allg_i.get(n)
        else:
            I = curr.get(n)
            ph = 3 if n == "-W1.4" else 1
        vals = [a, n, f, t, typ, cs.replace("x", "×") if cs != "—" else "—", s, L if L else None, I,
                ph if s else None]
        for k, v in enumerate(vals, start=1):
            c = ws.cell(r, k, v)
            c.font = F_LINK if (k == 9 and isinstance(v, str)) else F_N
            c.border = BOX
        ws.cell(r, 8).fill = YEL
        ws.cell(r, 8).font = F_IN
        if s:
            du = (f"=IF(J{r}=3,SQRT(3)*I{r}*({rho}*H{r}/G{r}*{cos}+{xl}*H{r}*SQRT(1-{cos}^2)),"
                  f"2*I{r}*{rho}*H{r}/G{r})")
            ws.cell(r, 11, du).number_format = "0.00"
            ws.cell(r, 12, f"=K{r}/IF(J{r}=3,{un},{u0})").number_format = "0.00%"
            if a.startswith("WV"):
                ws.cell(r, 13, f"=L{r}+Auslegung!$B${INP['R_DUWT']}").number_format = "0.00%"
            elif n == "-W0":
                ws.cell(r, 13, "vor Zähler")
            else:
                ws.cell(r, 13, f"=L{r}").number_format = "0.00%"
        for k in (11, 12, 13):
            ws.cell(r, k).font = F_N
            ws.cell(r, k).border = BOX
        ws.cell(r, 14, fn).font = F_N
        ws.cell(r, 14).border = BOX
        r += 1
    ws.cell(r + 1, 1, "Empfehlung OVE E 8101 Anh. G: ΔU ab Speisepunkt ≤ 4 %. Werte > 4 % in Spalte M → Querschnitt erhöhen "
                      "oder Länge kürzen.").font = F_G
    ws.freeze_panes = "C4"
    ws.auto_filter.ref = f"A{hr}:N{r - 1}"


def build(out):
    wb = Workbook()
    sheet_info(wb)
    sheet_ausl(wb)
    sheet_bom(wb)
    sheet_cables(wb)
    wb.save(out)


if __name__ == "__main__":
    build(sys.argv[1] if len(sys.argv) > 1 else "stueckliste.xlsx")
    print("ok")
