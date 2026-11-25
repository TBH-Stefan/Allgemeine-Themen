# -*- coding: utf-8 -*-
"""Gemeinsame Daten für Stromlaufplan (PDF) und Stückliste/Auslegung (XLSX).

Stockwerksverteiler (SV) mit 6 Zählerplätzen (Top 1-4, Allgemein, Reserve) und
4 Wohnungsverteiler (WV Top 1-4), Wien / Wiener Netze, Komponenten Schrack.
"""
import math

# ---------------------------------------------------------------- Projekt
DATE = "25.11.2026"
COMPANY = "TBH GmbH"
BEARB = "TBH"
PLAN_NO = "TBH-EP-2026-SV4WE"
PROJECT = "Wohnhaus Wien – 4 Wohneinheiten (Top 1–4)"
SUBTITLE = "Stockwerksverteiler mit Zählerplätzen + Wohnungsverteiler Top 1–4"
NETZ = "Wiener Netze GmbH"
TOPS = [1, 2, 3, 4]

# ---------------------------------------------------------------- Artikel
# status: V = verifiziert (Schrack-Datenblatt/Shopseite gelesen)
#         K = Schrack-Nr. aus Schrack-Katalog/Zubehörliste/Konfiguration (ohne Datenblatt)
#         O = Option / Abstimmung nötig     A = Art.-Nr. bei Schrack anfragen (Meterware)
#         B = Beistellung Netzbetreiber
ART = {
    # ---- Stockwerksverteiler
    "IL166333WS": dict(t="Unterputz-Zählerverteiler 3U-33G/WIEN, 6 Zählerplätze (2 Reihen à 3), "
                         "3 Nachzählerreihen (99 TE), H1605×B810×T250 mm, IP20C, Icc 25 kA, RAL 9016, "
                         "Drehriegel mit Doppelbart", m="Schrack", pv=0.0, st="V",
                       note="Komplettsatz Wien inkl. Zählertafeln nach OVE E 8640:2022; Pv,max 455 W"),
    "IL900189--": dict(t="Zählersteckklemme 50 A, 3-polig + 2N (Smart Meter Terminal), OVE R 21:2019, für Wien",
                       m="Schrack", pv=3.3, st="V", note="Pv 3,3 W bei 50 A"),
    "IL900191--": dict(t="Abdeckung für Zählersteckklemme (plombierbar)", m="Schrack", pv=0.0, st="V"),
    "IL900190--": dict(t="Steckstifte 63 A, 44 mm, VE = 7 Stk. (Zählerseite)", m="Schrack", pv=0.0, st="O",
                       note="nur wenn von Wiener Netze nicht beigestellt"),
    "SMART-METER": dict(t="Smart Meter Drehstrom, direkt angeschlossen (Zählertafel, 3-Punkt)",
                        m="Wiener Netze", pv=2.0, st="B", note="Beistellung Netzbetreiber; Pv Richtwert"),
    "IS503031--": dict(t="D02-Sicherungslasttrennschalter TYTAN T, 3-polig, 63 A, mit HK und "
                         "LED-Sicherungsausfallmeldung, plombierbar", m="Schrack", pv=5.4, st="V"),
    "IS504037--": dict(t="NEOZED-Sicherungseinsatz D02 35 A, gG", m="Schrack", pv=3.5, st="K",
                       note="Pv Richtwert je Einsatz"),
    "IS504036--": dict(t="NEOZED-Sicherungseinsatz D02 25 A, gG", m="Schrack", pv=2.6, st="K",
                       note="Pv Richtwert je Einsatz"),
    "IS504136--": dict(t="D02-Passhülse 35 A", m="Schrack", pv=0.0, st="K"),
    "IS504135--": dict(t="D02-Passhülse 25 A", m="Schrack", pv=0.0, st="K"),
    "IKB04035--": dict(t="Hauptabzweigklemme 4-polig 125 A, Eingang 1×35 mm², Abgänge 10×16 mm²",
                       m="Schrack", pv=2.4, st="K", note="Pv Richtwert"),
    "SI022270--": dict(t="Hauptklemme blank 25 mm², 1-polig, 1×6×16 mm² (PE-Verteilung)", m="Schrack",
                       pv=0.0, st="K"),
    "BZ900243--": dict(t="Hauptschalter (Lasttrennschalter) 40 A, 3-polig", m="Schrack", pv=2.2, st="V"),
    "IP3246353-": dict(t="Durchgangsklemme TB 16 I, 6–16 mm², grau", m="Phoenix Contact (über Schrack)",
                       pv=2.43, st="V"),
    "IP3076659-": dict(t="Durchgangsklemme TB 16 I BU, 6–16 mm², blau (N, mit Werkzeug trennbar)",
                       m="Phoenix Contact (über Schrack)", pv=2.43, st="V"),
    "IP3059896-": dict(t="Schutzleiterklemme TB 16-PE I, 6–16 mm², grün-gelb",
                       m="Phoenix Contact (über Schrack)", pv=0.0, st="V"),
    "IP3246324-": dict(t="Durchgangsklemme TB 4 I, 0,5–6 mm², grau", m="Phoenix Contact (über Schrack)",
                       pv=1.02, st="V"),
    "IP3057678-": dict(t="Durchgangsklemme TB 4 I BU, 0,5–6 mm², blau", m="Phoenix Contact (über Schrack)",
                       pv=1.02, st="V"),
    "IP3059883-": dict(t="Schutzleiterklemme TB 10-PE I, 1,5–16 mm², grün-gelb",
                       m="Phoenix Contact (über Schrack)", pv=0.0, st="V"),
    "AR634103-A": dict(t="FI-Schutzschalter 40 A, 4-polig, 30 mA, Typ A, G (kurzzeitverzögert), AMPARO",
                       m="Schrack", pv=10.0, st="V"),
    "AM018106--": dict(t="LS-Schalter AMPARO 10 kA, B 6 A, 1-polig", m="Schrack", pv=2.3, st="V"),
    "AM018110--": dict(t="LS-Schalter AMPARO 10 kA, B 10 A, 1-polig", m="Schrack", pv=2.3, st="V"),
    "AM018116--": dict(t="LS-Schalter AMPARO 10 kA, B 16 A, 1-polig", m="Schrack", pv=4.03, st="V"),
    "AM018316--": dict(t="LS-Schalter AMPARO 10 kA, B 16 A, 3-polig", m="Schrack", pv=12.08, st="V"),
    "BZ327360--": dict(t="Treppenlichtzeitschalter VOWA+, 1–12 min, 16 A, Ausschaltvorwarnung, Steuereingang",
                       m="Schrack", pv=1.0, st="V"),
    "BK002401": dict(t="Bezeichnungsstreifen 1000×10 mm, selbstklebend", m="Schrack", pv=0.0, st="V"),
    "IL900251": dict(t="Leerplatzabdeckung 45 mm, 1 m, kürzbar", m="Schrack", pv=0.0, st="V"),
    "IL900200-F": dict(t="Schließzylinder 61005 mit 1 Schlüssel (statt Doppelbart)", m="Schrack", pv=0.0,
                       st="O", note="nur bei Vorgabe Hausverwaltung"),
    # ---- Wohnungsverteiler
    "BK0850042-": dict(t="UP-Kleinverteiler-Wanne Massivwand, 4-reihig, 56 TE (4×14), IP40, inkl. N-/PE-Klemmen, "
                         "H717×B346×T94 mm", m="Schrack", pv=0.0, st="V"),
    "BK0851042F": dict(t="Rahmen, Stahlblechtür und Einsatz für BK085, 4-reihig, RAL 9003, IK08", m="Schrack",
                       pv=0.0, st="V", note="Pv,max 54 W (Datenblatt)"),
    "IS111314--": dict(t="Überspannungsableiter-Set 3+1 TT/TN-S, Typ 2 (+T3), 275 V, Imax 20 kA, steckbar",
                       m="Schrack", pv=0.08, st="V"),
    "BO218616--": dict(t="FI/LS-Schalter B16, 1+N, 30 mA, Typ A, G, 10 kA (BOLF)", m="Schrack", pv=3.2, st="V"),
    "BSB90116ME": dict(t="Phasenschiene Stift 3-polig 10 mm², 1 m (57 TE), kürzbar, für AMPARO", m="Schrack",
                       pv=21.5, st="V", note="Pv 21,5 W/m bei Nennstrom"),
    # ---- Meterware / Kleinmaterial (Art.-Nr. lt. Schrack-Kabelsortiment anfragen)
    "NHXHX-J 5x25": dict(t="Kabel NHXHX-J 5×25 mm² RM, halogenfrei", m="lt. Schrack-Kabelsortiment", pv=0.0,
                         st="A"),
    "NHXMH-J 5x10": dict(t="Mantelleitung NHXMH-J 5×10 mm², halogenfrei", m="lt. Schrack-Kabelsortiment",
                         pv=0.0, st="A"),
    "NHXMH-J 5x1,5": dict(t="Mantelleitung NHXMH-J 5×1,5 mm², halogenfrei", m="lt. Schrack-Kabelsortiment",
                          pv=0.0, st="A"),
    "NHXMH-J 3x1,5": dict(t="Mantelleitung NHXMH-J 3×1,5 mm², halogenfrei", m="lt. Schrack-Kabelsortiment",
                          pv=0.0, st="A"),
    "NYM-J 3x1,5": dict(t="Mantelleitung NYM-J 3×1,5 mm²", m="lt. Schrack-Kabelsortiment", pv=0.0, st="A"),
    "NYM-J 3x2,5": dict(t="Mantelleitung NYM-J 3×2,5 mm²", m="lt. Schrack-Kabelsortiment", pv=0.0, st="A"),
    "NYM-J 5x2,5": dict(t="Mantelleitung NYM-J 5×2,5 mm²", m="lt. Schrack-Kabelsortiment", pv=0.0, st="A"),
    "H07V-K 10 sw": dict(t="Aderleitung H07V-K 10 mm² schwarz (Verdrahtung)", m="lt. Schrack-Sortiment",
                         pv=0.0, st="A"),
    "H07V-K 10 bl": dict(t="Aderleitung H07V-K 10 mm² hellblau (Verdrahtung)", m="lt. Schrack-Sortiment",
                         pv=0.0, st="A"),
    "H07V-K 10 gnge": dict(t="Aderleitung H07V-K 10 mm² grün-gelb (Verdrahtung)", m="lt. Schrack-Sortiment",
                           pv=0.0, st="A"),
    "H07V-K 6": dict(t="Aderleitung H07V-K 6 mm² sw/bl/gnge (Ableiteranschluss ≤ 0,5 m)",
                     m="lt. Schrack-Sortiment", pv=0.0, st="A"),
    "H07V-K 1,5": dict(t="Aderleitung H07V-K 1,5 mm² sw/bl (Allgemein-Verdrahtung)", m="lt. Schrack-Sortiment",
                       pv=0.0, st="A"),
    "KLEINMAT": dict(t="Aderendhülsen, Endhalter/Abschlussdeckel TB-Klemmen, Kabelbinder, Beschriftung",
                     m="diverse", pv=0.0, st="A"),
}

STATUS_TXT = {"V": "verifiziert (Schrack-Datenblatt/Shop)",
              "K": "Schrack-Nr. lt. Katalog/Zubehörliste",
              "O": "Option – Abstimmung",
              "A": "Art.-Nr. bei Schrack anfragen",
              "B": "Beistellung Wiener Netze"}

# ---------------------------------------------------------------- Zählerplätze SV
ZP = [
    dict(n=1, name="Top 1", desc="Wohnung 3 Zi. (~75 m²)", fuse=35, fitted=True, kind="top", top=1),
    dict(n=2, name="Top 2", desc="Wohnung 3 Zi. (~75 m²)", fuse=35, fitted=True, kind="top", top=2),
    dict(n=3, name="Top 3", desc="Wohnung 3 Zi. (~75 m²)", fuse=35, fitted=True, kind="top", top=3),
    dict(n=4, name="Top 4", desc="Wohnung 3 Zi. (~75 m²)", fuse=35, fitted=True, kind="top", top=4),
    dict(n=5, name="Allgemein", desc="Stiegenhaus, Sprechanlage, RWA", fuse=25, fitted=True, kind="allg"),
    dict(n=6, name="Reserve", desc="vorbereitet (Wallbox/WP/PV)", fuse=35, fitted=False, kind="res"),
]

# ---------------------------------------------------------------- Allgemein-Abgänge (Zählerplatz 5)
# (bmk, LS-Art, char, In, Phase, vor_FI, Kabel, Typ, Länge m, Funktion, Klemmen)
ALLG = [
    dict(bmk="-5F3", art="AM018110--", ch="B", In=10, ph="L1", preFI=False, cable="-W5.1",
         typ="NHXMH-J 5x1,5", L=25, f="Stiegenhausbeleuchtung (über -5K1) + Sicherheitsleuchten (Dauerphase) + Taster",
         p=0.6),
    dict(bmk="-5F4", art="AM018106--", ch="B", In=6, ph="L2", preFI=False, cable="-W5.2",
         typ="NHXMH-J 3x1,5", L=15, f="Sprechanlage – Netzgerät", p=0.1),
    dict(bmk="-5F5", art="AM018116--", ch="B", In=16, ph="L3", preFI=False, cable=None,
         typ=None, L=0, f="Reserve", p=0.0),
    dict(bmk="-5F6", art="AM018110--", ch="B", In=10, ph="L3", preFI=True, cable="-W5.3",
         typ="NHXMH-J 3x1,5", L=20, f="RWA-Zentrale Stiegenhaus (ohne FI, gekennzeichnet)", p=0.3),
]

# ---------------------------------------------------------------- Wohnungsverteiler-Stromkreise
# Reihenfolge = Einbaureihenfolge auf Phasenschiene (L1,L2,L3,L1,...)
def C(bmk, f, In, poles, ph, typ, L, art, p_kw, short, cons):
    return dict(bmk=bmk, f=f, In=In, poles=poles, ph=ph, typ=typ, L=L, art=art, p=p_kw, short=short, cons=cons)

WV_G1 = [
    C("-F1.1", "Licht Vorraum / Wohnen / Küche", 10, 1, "L1", "NYM-J 3x1,5", 20, "AM018110--", 0.3, "Licht 1", "licht"),
    C("-F1.2", "Steckdosen Wohnzimmer", 16, 1, "L2", "NYM-J 3x2,5", 18, "AM018116--", 1.0, "SD Wohnen", "sd"),
    C("-F1.3", "Steckdosen Vorraum / WC / Abstellr.", 16, 1, "L3", "NYM-J 3x2,5", 18, "AM018116--", 0.5,
      "SD Vorraum", "sd"),
    C("-F1.4", "Kochfeld (Herdanschlussdose 3~)", 16, 3, "L1-L3", "NYM-J 5x2,5", 8, "AM018316--", 7.4,
      "Kochfeld", "geraet"),
    C("-F1.5", "Küche Arbeitsfläche 1", 16, 1, "L1", "NYM-J 3x2,5", 10, "AM018116--", 2.0, "Küche AF 1", "sd"),
    C("-F1.6", "Geschirrspüler", 16, 1, "L2", "NYM-J 3x2,5", 10, "AM018116--", 2.2, "Geschirrsp.", "geraet"),
    C("-F1.7", "Waschmaschine", 16, 1, "L3", "NYM-J 3x2,5", 12, "AM018116--", 2.2, "Waschmasch.", "geraet"),
    C("-F1.8", "Reserve", 16, 1, "L1", None, 0, "AM018116--", 0.0, "Reserve", "res"),
]
WV_G2 = [
    C("-F2.1", "Steckdosen Schlafzimmer", 16, 1, "L1", "NYM-J 3x2,5", 18, "AM018116--", 0.5, "SD Schlafen", "sd"),
    C("-F2.2", "Licht Schlafen / Kind / Bad / WC", 10, 1, "L2", "NYM-J 3x1,5", 20, "AM018110--", 0.3, "Licht 2",
      "licht"),
    C("-F2.3", "Licht Loggia / Abstellraum", 10, 1, "L3", "NYM-J 3x1,5", 15, "AM018110--", 0.2, "Licht 3", "licht"),
    C("-F2.4", "Backofen", 16, 1, "L1", "NYM-J 3x2,5", 8, "AM018116--", 3.0, "Backofen", "geraet"),
    C("-F2.5", "Steckdosen Kinderzimmer", 16, 1, "L2", "NYM-J 3x2,5", 18, "AM018116--", 0.5, "SD Kinder", "sd"),
    C("-F2.6", "Steckdosen Bad", 16, 1, "L3", "NYM-J 3x2,5", 12, "AM018116--", 1.5, "SD Bad", "sd"),
    C("-F2.7", "Fernwärme-Wohnungsstation / Raumthermostat", 10, 1, "L1", "NYM-J 3x1,5", 12, "AM018110--", 0.1,
      "FW-Station", "geraet"),
    C("-F2.8", "Küche Arbeitsfläche 2 / Mikrowelle", 16, 1, "L2", "NYM-J 3x2,5", 10, "AM018116--", 1.5,
      "Küche AF 2", "sd"),
    C("-F2.9", "Wäschetrockner", 16, 1, "L3", "NYM-J 3x2,5", 12, "AM018116--", 2.5, "Trockner", "geraet"),
    C("-F2.10", "Multimedia / Router", 16, 1, "L1", "NYM-J 3x2,5", 10, "AM018116--", 0.2, "Multimedia", "sd"),
    C("-F2.11", "Reserve", 16, 1, "L2", None, 0, "AM018116--", 0.0, "Reserve", "res"),
]
WV_F3 = C("-F3", "Kühl-/Gefriergerät (eigener FI/LS)", 16, 2, "L2", "NYM-J 3x2,5", 10, "BO218616--", 0.2,
          "Kühlgerät", "geraet")

# ---------------------------------------------------------------- Auslegung / Nachweise
CU_RHO = 0.0225      # Ohm*mm²/m, 1,25*rho20 (Betriebstemperatur), IEC 60364-5-52 Anh. G
X_PER_M = 0.08e-3    # Ohm/m
U_N = 400.0
U_0 = 230.0
COSPHI = 0.95

INPUT = dict(P_WE=14.5, g_WE=0.60, P_ALLG=2.0, P_RES=11.0, g_RES=1.0, L_W0=30.0, L_WTOP=25.0,
             RDF_SV=0.6, RDF_WV=0.5, Z_WV=0.35, IK_SV=10.0, L_BUS=0.40)


def du_3ph(I, L, S, cosphi=COSPHI):
    R = CU_RHO * L / S
    X = X_PER_M * L
    sinphi = math.sqrt(1 - cosphi ** 2)
    du = math.sqrt(3) * I * (R * cosphi + X * sinphi)
    return du, du / U_N * 100


def du_1ph(I, L, S, cosphi=1.0):
    R = CU_RHO * L / S
    X = X_PER_M * L
    sinphi = math.sqrt(max(0.0, 1 - cosphi ** 2))
    du = 2 * I * (R * cosphi + X * sinphi)
    return du, du / U_0 * 100


def calc():
    i = INPUT
    P = 4 * i["P_WE"] * i["g_WE"] + i["P_ALLG"] + i["P_RES"] * i["g_RES"]
    Ib = P * 1000 / (math.sqrt(3) * U_N * COSPHI)
    du_w0 = du_3ph(Ib, i["L_W0"], 25)
    du_wtop = du_3ph(35, i["L_WTOP"], 10)
    du_final = du_1ph(16, 20, 2.5)
    res = dict(P=P, Ib=Ib, du_w0=du_w0, du_wtop=du_wtop, du_final=du_final)
    # max. Länge W0 für 0,5 % und 1 %
    k = du_w0[1] / i["L_W0"]
    res["L_w0_05"] = 0.5 / k
    res["L_w0_10"] = 1.0 / k
    # Abschaltbedingung / max. Leitungslängen Endstromkreise (Z_WV = Schleifenimpedanz am WV, Annahme)
    lim = []
    for name, In, S, ph in (("B10 · NYM-J 3×1,5", 10, 1.5, 1), ("B16 · NYM-J 3×2,5", 16, 2.5, 1),
                            ("B16 3-pol. · NYM-J 5×2,5", 16, 2.5, 3)):
        zs = U_0 / (5 * In)
        l_abs = (zs - i["Z_WV"]) / (2 * CU_RHO / S)
        if ph == 1:
            l_du = 0.03 * U_0 / (2 * In * CU_RHO / S)
        else:
            l_du = 0.03 * U_N / (math.sqrt(3) * In * CU_RHO / S)
        lim.append((name, zs, l_abs, l_du))
    res["lim"] = lim
    # Kurzschluss: Ik3 am SV (Annahme 10 kA) -> am WV über 25 m 10 mm² (20 °C)
    z_src = U_0 / (i["IK_SV"] * 1000)
    r_w = 0.01786 * i["L_WTOP"] / 10
    res["ik_wv"] = U_0 / math.hypot(z_src + r_w, X_PER_M * i["L_WTOP"]) / 1000
    return res


def pv_wv():
    """Verlustleistung Wohnungsverteiler bei Nennströmen und mit RDF."""
    items = [("AR634103-A", 2), ("BO218616--", 1), ("IS111314--", 1), ("IKB04035--", 1),
             ("AM018110--", sum(1 for c in WV_G1 + WV_G2 if c["art"] == "AM018110--")),
             ("AM018116--", sum(1 for c in WV_G1 + WV_G2 if c["art"] == "AM018116--")),
             ("AM018316--", 1)]
    p = sum(ART[a]["pv"] * n for a, n in items) + ART["BSB90116ME"]["pv"] * INPUT["L_BUS"]
    return p, p * INPUT["RDF_WV"] ** 2


def pv_sv():
    items = [("IS503031--", 6), ("IS504037--", 12), ("IS504036--", 3), ("IL900189--", 6), ("SMART-METER", 5),
             ("BZ900243--", 6), ("AR634103-A", 1), ("AM018110--", 2), ("AM018106--", 1), ("AM018116--", 1),
             ("BZ327360--", 1), ("IP3246353-", 15), ("IP3076659-", 5), ("IP3246324-", 6), ("IP3057678-", 4),
             ("IKB04035--", 1)]
    p = sum(ART[a]["pv"] * n for a, n in items) + 10.0  # + Verdrahtung (Richtwert)
    return p, p * INPUT["RDF_SV"] ** 2


# ---------------------------------------------------------------- Stückliste
def bom():
    """Liefert Liste von Positionen: (Anlage, BMK, Menge, Einheit, Art, Bemerkung)."""
    sv = []
    a = sv.append
    a(("SV", "+SV (Gehäuse)", 1, "Stk", "IL166333WS", "Zählereinsatz IL186314-B (2×3 Zählertafeln) enthalten"))
    a(("SV", "-1X1…-6X1", 6, "Stk", "IL900189--", "1 je Zählerplatz"))
    a(("SV", "-1X1…-6X1", 6, "Stk", "IL900191--", "plombierbar"))
    a(("SV", "-1P1…-5P1", 5, "Stk", "SMART-METER", "Zählerplatz 6 = Reserve ohne Zähler"))
    a(("SV", "-1F1…-6F1", 6, "Stk", "IS503031--", "Vorzählersicherung, plombierbar"))
    a(("SV", "-1F1…-4F1, -6F1", 15, "Stk", "IS504136--", "3 je Lasttrenner (Top 1–4, Reserve)"))
    a(("SV", "-1F1…-4F1", 12, "Stk", "IS504037--", "3 je Top"))
    a(("SV", "-5F1", 3, "Stk", "IS504135--", "Allgemein"))
    a(("SV", "-5F1", 3, "Stk", "IS504036--", "Allgemein"))
    a(("SV", "-X0", 1, "Stk", "IKB04035--", "L1, L2, L3, N – ungezählter Bereich"))
    a(("SV", "-X0:PE", 1, "Stk", "SI022270--", "PE-Verteilung"))
    a(("SV", "-1Q1…-6Q1", 6, "Stk", "BZ900243--", "je Zählerplatz nach dem Zähler"))
    a(("SV", "-1X2…-4X2, -6X2", 15, "Stk", "IP3246353-", "L1–L3 Abgang"))
    a(("SV", "-1X2…-4X2, -6X2", 5, "Stk", "IP3076659-", "N Abgang (trennbar)"))
    a(("SV", "-1X2…-4X2, -6X2", 5, "Stk", "IP3059896-", "PE Abgang"))
    a(("SV", "-5F2", 1, "Stk", "AR634103-A", "Allgemein"))
    a(("SV", "-5F3, -5F6", 2, "Stk", "AM018110--", "Stiegenhaus, RWA"))
    a(("SV", "-5F4", 1, "Stk", "AM018106--", "Sprechanlage"))
    a(("SV", "-5F5", 1, "Stk", "AM018116--", "Reserve"))
    a(("SV", "-5K1", 1, "Stk", "BZ327360--", "Stiegenhauslicht"))
    a(("SV", "-5X2", 6, "Stk", "IP3246324-", "L-Abgänge Allgemein"))
    a(("SV", "-5X2", 4, "Stk", "IP3057678-", "N-Abgänge Allgemein"))
    a(("SV", "-5X2", 4, "Stk", "IP3059883-", "PE-Abgänge Allgemein"))
    a(("SV", "Zubehör", 1, "Stk", "BK002401", "Beschriftung Top/Abgänge"))
    a(("SV", "Zubehör", 1, "Stk", "IL900251", "freie TE abdecken"))
    a(("SV", "Option", 1, "Stk", "IL900200-F", "Option"))
    a(("SV", "Option", 6, "VE", "IL900190--", "Option"))
    a(("SV", "Verdrahtung", 20, "m", "H07V-K 10 sw", "Richtwert; kurz- und erdschlusssicher"))
    a(("SV", "Verdrahtung", 8, "m", "H07V-K 10 bl", "Richtwert"))
    a(("SV", "Verdrahtung", 8, "m", "H07V-K 10 gnge", "Richtwert"))
    a(("SV", "Verdrahtung", 6, "m", "H07V-K 1,5", "Richtwert"))
    a(("SV", "Kleinmaterial", 1, "Ps", "KLEINMAT", ""))
    a(("SV", "-W0", INPUT["L_W0"], "m", "NHXHX-J 5x25", "HV → SV, Länge Annahme"))
    a(("SV", "-W1…-W4", 4 * INPUT["L_WTOP"], "m", "NHXMH-J 5x10", "SV → WV Top 1–4, je ≤ 25 m"))
    a(("SV", "-W5.1", 25, "m", "NHXMH-J 5x1,5", "Länge Annahme"))
    a(("SV", "-W5.2, -W5.3", 35, "m", "NHXMH-J 3x1,5", "Länge Annahme"))

    wv = []
    b = wv.append
    n_b10 = sum(1 for c in WV_G1 + WV_G2 if c["art"] == "AM018110--")
    n_b16 = sum(1 for c in WV_G1 + WV_G2 if c["art"] == "AM018116--")
    l15 = sum(c["L"] for c in WV_G1 + WV_G2 if c["typ"] == "NYM-J 3x1,5")
    l25 = sum(c["L"] for c in WV_G1 + WV_G2 if c["typ"] == "NYM-J 3x2,5") + WV_F3["L"]
    l52 = sum(c["L"] for c in WV_G1 + WV_G2 if c["typ"] == "NYM-J 5x2,5")
    b(("WV", "+WV (Gehäuse)", 1, "Stk", "BK0850042-", "Unterputz-Wanne"))
    b(("WV", "+WV (Gehäuse)", 1, "Stk", "BK0851042F", "Rahmen/Tür/Einsatz"))
    b(("WV", "-X0", 1, "Stk", "IKB04035--", "Einspeisung L1–L3, N"))
    b(("WV", "-F01", 1, "Stk", "IS111314--", "vor den FI"))
    b(("WV", "-F1, -F2", 2, "Stk", "AR634103-A", "FI-Gruppe 1 und 2"))
    b(("WV", "-F3", 1, "Stk", "BO218616--", "Kühl-/Gefriergerät"))
    b(("WV", "-F1.1, -F2.2, -F2.3, -F2.7", n_b10, "Stk", "AM018110--", "Licht, FW-Station"))
    b(("WV", "-F1.2/.3/.5–.8, -F2.1/.4–.6/.8–.11", n_b16, "Stk", "AM018116--", "Steckdosen, Geräte, Reserve"))
    b(("WV", "-F1.4", 1, "Stk", "AM018316--", "Kochfeld"))
    b(("WV", "Phasenschiene", 1, "Stk", "BSB90116ME", "kürzen: 10 TE + 11 TE, Endkappen lt. Zubehör"))
    b(("WV", "Verdrahtung", 4, "m", "H07V-K 10 sw", "Richtwert"))
    b(("WV", "Verdrahtung", 2, "m", "H07V-K 10 bl", "Richtwert"))
    b(("WV", "Verdrahtung", 2, "m", "H07V-K 6", "Ableiter, Richtwert"))
    b(("WV", "Kleinmaterial", 1, "Ps", "KLEINMAT", ""))
    b(("WV", "-W1.x, -W2.x", l15, "m", "NYM-J 3x1,5", "Längen Annahme (Installationsplan)"))
    b(("WV", "-W1.x, -W2.x, -W3", l25, "m", "NYM-J 3x2,5", "Längen Annahme (Installationsplan)"))
    b(("WV", "-W1.4", l52, "m", "NYM-J 5x2,5", "Länge Annahme"))
    return sv, wv


# ---------------------------------------------------------------- Kabelliste
def cables():
    rows = []
    r = rows.append
    r(("SV", "-W0", "HV (NH00 gG 80 A)", "SV -X0", "NHXHX-J", "5x25", INPUT["L_W0"], "Anspeisung Stockwerksverteiler"))
    for t in TOPS:
        r(("SV", f"-W{t}", f"SV -{t}X2", f"WV{t} -X0", "NHXMH-J", "5x10", INPUT["L_WTOP"],
           f"Zuleitung Wohnungsverteiler Top {t}"))
    for c in ALLG:
        if c["cable"]:
            typ, cs = c["typ"].split(" ")
            r(("SV", c["cable"], "SV -5X2", "Allgemein", typ, cs, c["L"], c["f"]))
    r(("SV", "-W6", "SV -6X2", "Reserve", "—", "—", 0, "Reserve-Zählerplatz (nicht verlegt)"))
    for t in TOPS:
        for c in WV_G1 + WV_G2 + [WV_F3]:
            if c["typ"]:
                typ, cs = c["typ"].split(" ")
                r((f"WV{t}", c["bmk"].replace("-F", "-W"), f"WV{t} {c['bmk']}", c["short"], typ, cs, c["L"], c["f"]))
    return rows


if __name__ == "__main__":
    print(calc())
    print("Pv WV", pv_wv(), "Pv SV", pv_sv())
    sv, wv = bom()
    print(len(sv), len(wv), len(cables()))
