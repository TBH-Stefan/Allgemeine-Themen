"""Erzeugt Auslegung_Stueckliste_PV-Ueberschuss.xlsx (Auslegung, Stückliste, Klemmenplan, Parameter, IBN)."""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent / "Auslegung_Stueckliste_PV-Ueberschuss.xlsx"

F = "Arial"
f_norm = Font(name=F, size=10)
f_bold = Font(name=F, size=10, bold=True)
f_title = Font(name=F, size=14, bold=True)
f_h2 = Font(name=F, size=11, bold=True)
f_in = Font(name=F, size=10, color="0000FF")
f_link = Font(name=F, size=10, color="008000")
f_small = Font(name=F, size=9, italic=True, color="555555")
fill_in = PatternFill("solid", start_color="FFFF00")
fill_hdr = PatternFill("solid", start_color="D9E1F2")
fill_res = PatternFill("solid", start_color="E2EFDA")
thin = Side(style="thin", color="999999")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")

wb = Workbook()

# ---------------------------------------------------------------- Auslegung
ws = wb.active
ws.title = "Auslegung"
ws["A1"] = "Auslegung PV-Überschussnutzung SolarEdge SE8K → WOLF CHA-10 / WW-Speicher"
ws["A1"].font = f_title
ws["A2"] = "Stand 18.11.2026 · Index A · gehört zu Anschlussplan_PV-Ueberschuss_SE8K_WOLF-CHA10.pdf"
ws["A2"].font = f_small
ws["A3"] = "Legende: gelbe Zellen mit blauer Schrift = Eingabe/Annahme (anpassen) · schwarz = Formel · grün hinterlegt = Ergebnis"
ws["A3"].font = f_small

hdr = ["Nr", "Größe", "Wert", "Einheit", "Quelle / Annahme"]
r = 5
for c, h in enumerate(hdr, 1):
    cell = ws.cell(r, c, h)
    cell.font = f_bold
    cell.fill = fill_hdr
    cell.border = box

ref = {}  # Schlüssel -> Zelladresse Wert


def row_in(key, name, val, unit, src, fmt="0.0"):
    global r
    r += 1
    ws.cell(r, 1, len(ref) + 1).font = f_norm
    ws.cell(r, 2, name).font = f_norm
    v = ws.cell(r, 3, val)
    v.font = f_in
    v.fill = fill_in
    v.number_format = fmt
    ws.cell(r, 4, unit).font = f_norm
    ws.cell(r, 5, src).font = f_norm
    for c in range(1, 6):
        ws.cell(r, c).border = box
        ws.cell(r, c).alignment = wrap
    ref[key] = f"C{r}"


def row_f(key, name, formula, unit, note, fmt="0.00", result=False):
    global r
    r += 1
    ws.cell(r, 1, len(ref) + 1).font = f_norm
    ws.cell(r, 2, name).font = f_bold if result else f_norm
    v = ws.cell(r, 3, formula)
    v.font = f_bold if result else f_norm
    v.number_format = fmt
    if result:
        v.fill = fill_res
    ws.cell(r, 4, unit).font = f_norm
    ws.cell(r, 5, note).font = f_norm
    for c in range(1, 6):
        ws.cell(r, c).border = box
        ws.cell(r, c).alignment = wrap
    ref[key] = f"C{r}"


def section(title):
    global r
    r += 2
    ws.cell(r, 1, title).font = f_h2


section("Eingaben")
row_in("V", "Speichervolumen V", 300, "l", "Annahme – Typenschild Speicher prüfen", "0")
row_in("tkw", "Kaltwassertemperatur ϑ_KW", 10, "°C", "Annahme")
row_in("tsoll", "WW-Solltemperatur (BM-2)", 50, "°C", "Nutzerangabe")
row_in("wp027", "PV-Anhebung WW (WP027)", 10, "K", "Empfehlung; WOLF-Bereich 0–20 K (PU Wärmepumpe und Photovoltaik)")
row_in("tmax", "Abschalttemperatur Heizpatrone (TR)", 75, "°C", "Nutzerangabe")
row_in("a14", "WW-Maximaltemperatur WOLF (A14)", 65, "°C", "Annahme – am BM-2 ablesen")
row_in("tmix", "Mischwasser-Zapftemperatur", 40, "°C", "Annahme (Dusche)")
row_in("php", "Leistung Heizpatrone P_HP", 3.0, "kW", "Annahme – Typenschild −E1; −A2 max. 3,0 kW")
row_in("pwp", "el. Leistung CHA-10 im WW-Betrieb P_WP", 2.0, "kW", "Annahme – Datenblatt/Monitoring prüfen")
row_in("res", "Reserve Einschaltschwelle −A1", 0.3, "kW", "Annahme (Grundlastschwankung)")
row_in("copPV", "COP WW-Ladung 50 → 60 °C", 2.5, "–", "Annahme")
row_in("copN", "COP WW-Ladung normal (bis 50 °C)", 3.0, "–", "Annahme")
row_in("U", "Netzspannung 1~", 230, "V", "Nennwert", "0")
row_in("L", "Leitungslänge −W5 + −W6 (einfach)", 8, "m", "Annahme – aufmessen", "0")
row_in("A", "Querschnitt −W5/−W6", 2.5, "mm²", "Planung")
row_in("kappa", "Leitfähigkeit Cu κ", 56, "m/(Ω·mm²)", "Kupfer bei 20 °C", "0")
row_in("qbv", "Bereitschaftsverlust Speicher", 1.6, "kWh/24h", "Annahme – Datenblatt (Prüfung nach EN 12897)")
row_in("tref", "Bezugstemperatur Bereitschaftsverlust", 65, "°C", "Prüfbedingung Datenblatt", "0")
row_in("tu", "Aufstellraumtemperatur", 20, "°C", "Annahme / Prüfbedingung", "0")
row_in("kS", "Strombezugspreis", 0.25, "€/kWh", "Beispielwert – anpassen", "0.000")
row_in("kE", "Einspeisevergütung", 0.06, "€/kWh", "Beispielwert – anpassen", "0.000")
row_in("n", "Vollladungen 50 → 75 °C pro Jahr", 120, "1/a", "Annahme (sonnige Tage)", "0")
row_in("cw", "spez. Wärmekapazität Wasser c_w", 1.163, "Wh/(kg·K)", "Konstante; Dichte ≈ 1 kg/l (Näherung)", "0.000")

R = ref  # Kurzform
section("Berechnung")
row_f("t1", "Temperatur nach Stufe 1 (WP)", f"=MIN({R['tsoll']}+{R['wp027']},{R['a14']})", "°C",
      "WW-Soll + WP027, begrenzt durch A14", "0.0")
row_f("chk27", "Prüfung WP027 einstellbar", f'=IF({R["wp027"]}<=20,"OK","> 20 K nicht einstellbar")', "", "WOLF max. 20 K", "General")
row_f("chkT", "Prüfung TR über Stufe 1", f'=IF({R["tmax"]}>{ref["t1"]},"OK","TR ≤ Stufe 1 – Stufe 2 ohne Wirkung")', "", "", "General")
row_f("q1", "Wärme Stufe 1 Q1", f"={R['V']}*{R['cw']}*({ref['t1']}-{R['tsoll']})/1000", "kWh", "V·c·Δϑ")
row_f("q2", "Wärme Stufe 2 Q2 (Heizpatrone)", f"={R['V']}*{R['cw']}*MAX({R['tmax']}-{ref['t1']},0)/1000", "kWh", "V·c·Δϑ")
row_f("qg", "Zusätzlich gespeicherte Wärme Q_ges", f"={ref['q1']}+{ref['q2']}", "kWh", "über WW-Soll", result=True)
row_f("e1", "PV-Strom Stufe 1", f"={ref['q1']}/{R['copPV']}", "kWh", "Q1 / COP")
row_f("e2", "PV-Strom Stufe 2", f"={ref['q2']}", "kWh", "Wirkungsgrad Heizpatrone ≈ 1")
row_f("eg", "PV-Strom gesamt je Vollladung", f"={ref['e1']}+{ref['e2']}", "kWh", "", result=True)
row_f("tl1", "Ladezeit Stufe 1", f"=IF({R['pwp']}*{R['copPV']}>0,{ref['q1']}/({R['pwp']}*{R['copPV']}),0)", "h", "bei vollem Überschuss")
row_f("tl2", "Ladezeit Stufe 2", f"=IF({R['php']}>0,{ref['q2']}/{R['php']},0)", "h", "bei voller Leistung −E1")
row_f("s1", "Einschaltschwelle −A1 (SolarEdge)", f"={R['pwp']}+{R['res']}", "kW", "in SetApp eintragen", "0.0", result=True)
row_f("smax", "Überschuss für beide Stufen gleichzeitig", f"={R['pwp']}+{R['php']}", "kW", "darüber Einspeisung", "0.0")
row_f("I", "Strom Heizpatrone", f"={R['php']}*1000/{R['U']}", "A", "1~, ohmsch", "0.0")
row_f("chkF5", "Schutzorgan −F5", f'=IF({ref["I"]}<=16,"B16 ausreichend","> 16 A – Konzept ändern")', "", "FI/LS 1P+N B16 30 mA Typ A", "General")
row_f("chkA2", "Prüfung −A2 (max. 3,0 kW)", f'=IF({R["php"]}<=3,"OK – Home Hot Water Controller","> 3 kW: Option −A3/−Q1 oder AC•THOR")', "", "", "General")
row_f("du", "Spannungsfall −W5/−W6", f"=2*{R['L']}*{ref['I']}/({R['kappa']}*{R['A']})/{R['U']}", "%",
      "ΔU = 2·L·I/(κ·A); Richtwert ≤ 3 %", "0.00%")
row_f("vm1", "Mischwassermenge bei WW-Soll", f"={R['V']}*({R['tsoll']}-{R['tkw']})/({R['tmix']}-{R['tkw']})", "l", "", "0")
row_f("vm2", "Mischwassermenge bei Maximaltemperatur", f"={R['V']}*({R['tmax']}-{R['tkw']})/({R['tmix']}-{R['tkw']})", "l", "", "0")
row_f("dvm", "Mehr-Mischwasser je Vollladung", f"={ref['vm2']}-{ref['vm1']}", "l", "", "0", result=True)
row_f("qv", "Mehr-Bereitschaftsverlust (75 statt 50 °C, 24 h)", f"={R['qbv']}*({R['tmax']}-{R['tsoll']})/({R['tref']}-{R['tu']})", "kWh/24h",
      "linear über Temperaturdifferenz (Näherung)")
row_f("nu", "Nutzen je Vollladung (vermiedener Bezug)", f"=({ref['qg']}-{ref['qv']})/{R['copN']}*{R['kS']}", "€", "(Q_ges − Verlust)/COP_normal · Strompreis", "0.00")
row_f("ko", "Entgangene Einspeisevergütung", f"={ref['eg']}*{R['kE']}", "€", "", "0.00")
row_f("ne", "Netto je Vollladung", f"={ref['nu']}-{ref['ko']}", "€", "", "0.00", result=True)
row_f("na", "Netto pro Jahr", f"={ref['ne']}*{R['n']}", "€/a", "", "0", result=True)
row_f("neo", "Vergleich: Netto/Jahr nur Heizpatrone (ohne WP-Stufe)",
      f"=((({R['V']}*{R['cw']}*({R['tmax']}-{R['tsoll']})/1000)-{ref['qv']})/{R['copN']}*{R['kS']}"
      f"-({R['V']}*{R['cw']}*({R['tmax']}-{R['tsoll']})/1000)*{R['kE']})*{R['n']}", "€/a",
      "zeigt Vorteil der WP-Stufe", "0")

# Sensitivität
section("Sensitivität Netto pro Jahr [€/a] – Strombezugspreis (Zeilen) × Einspeisevergütung (Spalten)")
r += 1
ws.cell(r, 2, "Bezug \\ Einspeisung").font = f_bold
feeds = [0.03, 0.05, 0.07, 0.09]
prices = [0.20, 0.25, 0.30, 0.35]
hdr_row = r
for j, fe in enumerate(feeds):
    c = ws.cell(r, 3 + j, fe)
    c.font = f_in
    c.fill = fill_in
    c.number_format = "0.00"
    c.border = box
for i, pr in enumerate(prices):
    r += 1
    c = ws.cell(r, 2, pr)
    c.font = f_in
    c.fill = fill_in
    c.number_format = "0.00"
    c.border = box
    for j in range(len(feeds)):
        col = 3 + j
        colL = ws.cell(hdr_row, col).column_letter
        f = (f"=(({ref['qg']}-{ref['qv']})/{R['copN']}*$B{r}-{ref['eg']}*{colL}${hdr_row})*{R['n']}")
        cc = ws.cell(r, col, f)
        cc.number_format = "0"
        cc.font = f_norm
        cc.border = box

for col, w in zip("ABCDEFG", (5, 46, 16, 12, 58, 12, 12)):
    ws.column_dimensions[col].width = w
ws.freeze_panes = "A6"

# ---------------------------------------------------------------- Stückliste
sl = wb.create_sheet("Stückliste")
sl["A1"] = "Stückliste PV-Überschussnutzung SE8K → WOLF CHA-10"
sl["A1"].font = f_title
sl["A2"] = ("Legende: gelbe Zellen eintragen – Option „Ja/Nein“, EP netto in € (z. B. 123,45), Rabatt in % (z. B. 15 %). "
            "Keine Listenpreise hinterlegt. GP = Menge × EP × (1 − Rabatt), nur wenn Option = Ja.")
sl["A2"].font = f_small
cols = ["Pos", "BMK", "Bezeichnung", "Typ / Art.-Nr.", "Hersteller", "Menge", "Einh.", "Option", "EP netto €",
        "Rabatt", "GP netto €", "Bemerkung"]
for c, h in enumerate(cols, 1):
    cell = sl.cell(4, c, h)
    cell.font = f_bold
    cell.fill = fill_hdr
    cell.border = box
items = [
    ("−P1", "Energy Meter 3~ 65 A direkt, Modbus/Home Network", "MTR-240-3PC1-D-A-MW", "SolarEdge", 1, "Stk", "Ja",
     "Nein, falls Zähler (Export+Import) vorhanden"),
    ("−", "Home-Network-Plug-in für SE8K", "lt. SolarEdge (Typenschild prüfen)", "SolarEdge", 1, "Stk", "Nein",
     "nur falls WR nicht Home-Network-fähig"),
    ("−A1", "Home Load Controller (Relais NO 8 A)", "SEM-DCS-R08-00", "SolarEdge", 1, "Stk", "Ja", "Stufe 1 – WOLF PV-Kontakt"),
    ("−A2", "Home Hot Water Controller 3 kW", "SMRT-HOT-WTR-30-S2", "SolarEdge", 1, "Stk", "Ja",
     "Stufe 2 – Heizpatrone stufenlos, nur ohmsch ≤ 3 kW"),
    ("−K1", "Koppelrelais 230 V AC, 1 Wechsler, Hartgoldkontakt, Hutschiene", "z. B. Phoenix PLC-RSC…/21AU, Finder 38.51 Au",
     "nach Wahl", 1, "Stk", "Ja", "SELV-Trennung, Kleinsignal"),
    ("−F4", "FI/LS 1P+N B10 30 mA Typ A", "nach Hausstandard", "nach Wahl", 1, "Stk", "Ja", "Versorgung −A1/−K1"),
    ("−F5", "FI/LS 1P+N B16 30 mA Typ A", "nach Hausstandard", "nach Wahl", 1, "Stk", "Ja", "Versorgung −A2/−E1"),
    ("−E1", "E-Heizpatrone 1~ 230 V ≤ 3 kW mit TR (bis ≥ 75 °C) + STB", "passend zur Speichermuffe", "nach Wahl", 1, "Stk",
     "Nein", "Ja, falls keine geeignete vorhanden"),
    ("−TV1", "Thermostatischer Verbrühschutzmischer DN 20", "Einstellbereich bis 55 °C", "nach Wahl", 1, "Stk", "Ja",
     "zwingend bei 75 °C"),
    ("−W1", "J-Y(St)Y 2×2×0,8", "", "", 10, "m", "Ja", "entfällt bei Funk"),
    ("−W4", "J-Y(St)Y 2×2×0,8", "", "", 15, "m", "Ja", "SELV zu WOLF X0"),
    ("−W2/−W3", "NYM-J 3×1,5 / H07V-K 1,5", "", "", 3, "m", "Ja", ""),
    ("−W5/−W6", "NYM-J 3×2,5", "", "", 8, "m", "Ja", "letzter Abschnitt wärmebeständig"),
    ("", "Kleinmaterial (Aderendhülsen, Beschriftung, Kanal)", "", "", 1, "Psch", "Ja", ""),
    ("", "Montage Elektro", "", "", 6, "h", "Ja", "Annahme Stunden"),
    ("", "Inbetriebnahme SolarEdge + WOLF-Parametrierung", "", "", 2, "h", "Ja", "Annahme Stunden"),
    ("−A3", "Option 3~-Heizpatrone: 2. Home Load Controller", "SEM-DCS-R08-00", "SolarEdge", 1, "Stk", "Nein",
     "nur Ein/Aus"),
    ("−Q1", "Option 3~-Heizpatrone: Schütz 3-pol. 25 A, Spule 230 V", "nach Wahl", "nach Wahl", 1, "Stk", "Nein",
     "mit −A3"),
    ("−F6", "Option 3~-Heizpatrone: FI/LS 3P+N C16 30 mA Typ A", "nach Wahl", "nach Wahl", 1, "Stk", "Nein", "mit −A3"),
    ("−A2*", "Alternative zu −A2/−A3: Leistungssteller 3~ stufenlos bis 9 kW", "my-PV AC•THOR 9s", "my-PV", 1, "Stk",
     "Nein", "Modbus TCP am SE8K – nur 1 TCP-Client, Kompatibilität prüfen"),
]
dv = DataValidation(type="list", formula1='"Ja,Nein"', allow_blank=False)
sl.add_data_validation(dv)
row = 4
for i, it in enumerate(items, 1):
    row += 1
    bmk, name, typ, herst, qty, unit, opt, note = it
    vals = [i, bmk, name, typ, herst, qty, unit, opt, None, None,
            f'=IF(H{row}="Ja",F{row}*I{row}*(1-J{row}),0)', note]
    for c, v in enumerate(vals, 1):
        cell = sl.cell(row, c, v)
        cell.font = f_norm
        cell.border = box
        cell.alignment = wrap
    for c in (8, 9, 10):
        sl.cell(row, c).fill = fill_in
        sl.cell(row, c).font = f_in
    sl.cell(row, 9).number_format = "#,##0.00"
    sl.cell(row, 10).number_format = "0%"
    sl.cell(row, 11).number_format = "#,##0.00"
    dv.add(f"H{row}")
last = row
row += 1
sl.cell(row, 10, "Summe netto").font = f_bold
s = sl.cell(row, 11, f"=SUM(K5:K{last})")
s.font = f_bold
s.number_format = "#,##0.00"
s.fill = fill_res
s.border = box
for col, w in zip("ABCDEFGHIJKL", (5, 9, 46, 30, 12, 7, 6, 8, 11, 8, 12, 40)):
    sl.column_dimensions[col].width = w
sl.freeze_panes = "A5"

# ---------------------------------------------------------------- Klemmenplan
kp = wb.create_sheet("Klemmenplan")
kp["A1"] = "Klemmenplan"
kp["A1"].font = f_title
kh = ["Nr", "Von", "Nach", "Kabel / Ader", "Funktion", "Spannung"]
for c, h in enumerate(kh, 1):
    cell = kp.cell(3, c, h)
    cell.font = f_bold
    cell.fill = fill_hdr
    cell.border = box
krows = [
    ("−F4: L / N", "−A1: L / N", "−W2 · L, N, PE", "Versorgung Load Controller", "230 V AC"),
    ("−F4: L", "−A1: COM", "Brücke 1,5 mm²", "L auf Relaiskontakt", "230 V AC"),
    ("−A1: NO", "−K1: A1", "−W3", "Spule Koppelrelais", "230 V AC"),
    ("−K1: A2", "N-Klemme UV", "−W3 (blau)", "Rückleiter Spule", "230 V AC"),
    ("−K1: 11", "WOLF X0: PV/SG1", "−W4 · Paar 1 Ader a", "PV-Anhebung ein", "SELV (WOLF)"),
    ("−K1: 14", "WOLF X0: GND", "−W4 · Paar 1 Ader b", "Bezugspotential", "SELV (WOLF)"),
    ("Schirm −W4", "PE Verteilerseite", "–", "einseitig auflegen", "–"),
    ("WOLF X0: EVU/SG0", "WOLF X0: GND", "bestehend", "Brücke bzw. EVU-Kontakt – nicht ändern", "SELV"),
    ("−F5: L / N", "−A2: L-IN / N-IN / PE", "−W5 · L, N, PE", "Versorgung Hot Water Controller", "230 V AC"),
    ("−A2: L-OUT / N-OUT / PE", "−E1: L / N / PE (über TR + STB)", "−W6 · L, N, PE", "Heizpatrone geregelt 0–3 kW", "230 V AC"),
    ("SE8K RS485-1: A / B / G", "−P1: A / B / GND", "−W1 · Paar 1 = A/B, Paar 2 = G", "Modbus RTU, Zähler-ID 2", "SELV"),
    ("Schirm −W1", "PE WR-Seite", "–", "einseitig auflegen", "–"),
    ("SE8K RS485-1 / −P1", "–", "–", "Abschluss 120 Ω an beiden Busenden ein", "–"),
    ("WW-Speicher", "Potentialausgleich", "bestehend prüfen", "Schutzpotentialausgleich", "–"),
]
for i, kr in enumerate(krows, 1):
    rr = 3 + i
    for c, v in enumerate((i,) + kr, 1):
        cell = kp.cell(rr, c, v)
        cell.font = f_norm
        cell.border = box
        cell.alignment = wrap
kp.cell(4 + len(krows) + 1, 1,
        "Klemmenbezeichnungen −A1/−A2/−P1/SE8K lt. aktueller Installationsanleitung; X0 lt. WOLF-Montageanleitung CHA.").font = f_small
for col, w in zip("ABCDEF", (5, 26, 30, 30, 40, 14)):
    kp.column_dimensions[col].width = w

# ---------------------------------------------------------------- Parameter
pa = wb.create_sheet("Parameter")
pa["A1"] = "Parametrierung WOLF BM-2 (FW 3.10) und SolarEdge"
pa["A1"].font = f_title
pa["A2"] = "Grün = Verknüpfung zur Auslegung. Werte vor Ort ablesen/protokollieren (Spalte Ist)."
pa["A2"].font = f_small
ph = ["System", "Parameter", "Bezeichnung", "Werk", "Soll", "Ist (vor Ort)", "Bemerkung"]
for c, h in enumerate(ph, 1):
    cell = pa.cell(4, c, h)
    cell.font = f_bold
    cell.fill = fill_hdr
    cell.border = box
A = "Auslegung!"
prow = [
    ("WOLF", "WP025", "Funktion Eingänge SG/PV", "je Gerät", "PV", "PV/SG1 = PV-Anhebung; EVU/SG0 bleibt EVU-Sperre"),
    ("WOLF", "WP026", "Externe Anhebung Heizung", "0 K", "0 K", "optional 1–2 K in der Heizperiode"),
    ("WOLF", "WP027", "Externe Anhebung WW [K]", "0 K", f"={A}{ref['wp027']}", "aus Auslegung; Bereich 0–20 K"),
    ("WOLF", "WP028", "WEZ-Management bei PV/SG", "Standard", "nur WP", "WOLF-E-Heizelement nicht zuschalten"),
    ("WOLF", "WP032", "Heizen bei PV/SG", "Ein", "Ein", "unverändert"),
    ("WOLF", "WP033", "Kühlen bei PV/SG", "Aus", "Aus", "unverändert"),
    ("WOLF", "WP034–036", "Bivalenzpunkte bei PV/SG", "−25/−5/−25 °C", "unverändert", ""),
    ("WOLF", "A14", "WW-Maximaltemperatur [°C]", "Anlage", f"={A}{ref['t1']}", "mind. WW-Soll + WP027; nicht 75 °C"),
    ("WOLF", "–", "WW-Solltemperatur [°C]", "–", f"={A}{ref['tsoll']}", "Benutzerebene, unverändert"),
    ("SolarEdge", "−P1", "Zählerfunktion / Ort", "–", "Export+Import", "Netzanschlusspunkt"),
    ("SolarEdge", "−P1", "Kommunikation", "–", "RS485-1 Modbus, ID 2", "oder Home Network (Funk)"),
    ("SolarEdge", "−A1", "Betriebsart", "–", "SG Ready 1-stufig", "alternativ Relais „Überschuss“"),
    ("SolarEdge", "−A1", "Aktivierungsleistung [kW]", "–", f"={A}{ref['s1']}", "aus Auslegung"),
    ("SolarEdge", "−A1", "Mindest-Einschaltdauer", "–", "30 min", "Takten vermeiden"),
    ("SolarEdge", "−A1", "Zeitfenster", "–", "optional 10–16 Uhr", ""),
    ("SolarEdge", "−A1/−A2", "Priorität", "–", "1 / 2", "WP vor Heizpatrone"),
    ("SolarEdge", "−A2", "Nennleistung Heizpatrone [kW]", "–", f"={A}{ref['php']}", "Typenschild −E1"),
    ("SolarEdge", "−A2", "Modus", "–", "PV-Überschuss automatisch", "kein Zeit-Boost aus dem Netz"),
]
for i, p in enumerate(prow, 1):
    rr = 4 + i
    sysn, par, bez, werk, soll, bem = p
    for c, v in enumerate((sysn, par, bez, werk, soll, None, bem), 1):
        cell = pa.cell(rr, c, v)
        cell.font = f_norm
        cell.border = box
        cell.alignment = wrap
    if isinstance(soll, str) and soll.startswith("="):
        pa.cell(rr, 5).font = f_link
        pa.cell(rr, 5).number_format = "0.0"
    pa.cell(rr, 6).fill = fill_in
    pa.cell(rr, 6).font = f_in
pa.cell(4 + len(prow) + 2, 1,
        "Quelle WOLF-Parameter: Planungsunterlage „Wärmepumpe und Photovoltaik“ (4801950, 12/2023), Betriebsanleitung CHA "
        "für die Fachkraft. Bezeichnungen können je nach Softwarestand BM-2/HCM-4 abweichen.").font = f_small
for col, w in zip("ABCDEFG", (11, 11, 32, 15, 24, 16, 44)):
    pa.column_dimensions[col].width = w
pa["E7"].comment = Comment("Verknüpft mit Auslegung (WP027).", "Auslegung")

# ---------------------------------------------------------------- IBN
ib = wb.create_sheet("IBN")
ib["A1"] = "Inbetriebnahme-Checkliste"
ib["A1"].font = f_title
ih = ["Nr", "Prüfschritt", "Sollergebnis", "i.O. (ja/nein)", "Messwert / Bemerkung", "Datum / Visum"]
for c, h in enumerate(ih, 1):
    cell = ib.cell(3, c, h)
    cell.font = f_bold
    cell.fill = fill_hdr
    cell.border = box
checks = [
    ("Erstprüfung Stromkreise −F4/−F5 (OVE E 8101 Teil 6)", "R_iso, Z_S, RCD-Auslösung i.O."),
    ("−P1 in SetApp sichtbar, Vorzeichen prüfen", "Einspeisung = Export, Bezug = Import"),
    ("−A1/−A2 im Home Network gekoppelt", "online"),
    ("−A1 manuell EIN (Taste)", "−K1 zieht an; BM-2: PV-Anhebung aktiv, WW-Soll 60 °C"),
    ("−A2 manuell EIN", "Strom −E1 ≈ 13 A (Zangenmessung)"),
    ("TR −E1 auf 75 °C, STB unverstellt", "bei 75 °C → 0 A"),
    ("Mischwasser an nächster Zapfstelle", "≤ 55 °C"),
    ("WOLF nach erster 75-°C-Ladung", "keine Stör-/Warnmeldung"),
    ("Automatikbetrieb an sonnigem Tag", "Zustände 0 → 1 → 2 → 3 lt. Anschlussplan Blatt 1"),
    ("Parameter- und Messprotokoll", "im Anlagenbuch abgelegt"),
]
for i, (a, b) in enumerate(checks, 1):
    rr = 3 + i
    for c, v in enumerate((i, a, b, None, None, None), 1):
        cell = ib.cell(rr, c, v)
        cell.font = f_norm
        cell.border = box
        cell.alignment = wrap
    for c in (4, 5, 6):
        ib.cell(rr, c).fill = fill_in
for col, w in zip("ABCDEF", (5, 44, 48, 13, 30, 16)):
    ib.column_dimensions[col].width = w

for sheet in wb.worksheets:
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print(OUT)
print({k: v for k, v in ref.items() if k in ("t1", "q1", "q2", "qg", "eg", "s1", "I", "du", "na", "neo")})
