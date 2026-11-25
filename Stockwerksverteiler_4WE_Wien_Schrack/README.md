# Stockwerksverteiler mit Zählerplätzen + Wohnungsverteiler Top 1–4 (Wien, Schrack)

Allpoliger Stromlaufplan (Mehrlinienstromlaufplan) und Stückliste für einen Stockwerksverteiler mit 6 Zählerplätzen (Top 1–4, Allgemein, Reserve) im Netz der Wiener Netze sowie vier typengleiche Wohnungsverteiler. Komponenten: Schrack Technik (Eigenmarke) und über Schrack vertriebene Unterlieferanten (Phoenix Contact). Der Planaufbau orientiert sich an vorliegenden EPLAN-Plänen der Fa. Csernohorszky: Deckblatt mit IEC-61439-Daten, Inhalt, Ansicht, Anspeisung, Abgänge, Kabelübersicht, Stückliste mit Verlustleistung. Übernommen wurde nur die Struktur, keine Inhalte.

| Datei | Inhalt |
|---|---|
| `Stromlaufplan_SV-WV_Wien_4WE.pdf` | 33 Blätter A3 quer: Deckblatt, Inhalt, Übersichtsschaltplan (Energiefluss), Auslegung/Nachweise, SV-Ansicht, Anspeisung, Zählerplätze, Allgemein-Abgänge, je Top: Ansicht, Anspeisung + Ableiter + Kühlgerät, FI-Gruppe 1, FI-Gruppe 2 (2 Blätter), Kabelübersicht, Stückliste |
| `Stueckliste_Auslegung_SV-WV_Wien_4WE.xlsx` | Info/Legende, Auslegung mit Formeln (Leistung, Ib, Leiterschutz, ΔU, Selektivität, Ik, Verlustleistung, max. Leitungslängen, Sensitivitäten), Stückliste mit Preis-/Rabatt-/Optionsfeldern, Kabelliste mit ΔU je Leitung |
| `src/` | `data.py` (gemeinsame Daten), `plan.py` → PDF (`python3 src/plan.py <out.pdf>`), `build_xlsx.py` → Excel (`python3 src/build_xlsx.py <out.xlsx>`, danach mit LibreOffice neu berechnen) |

## Konzept (Vorgaben aus der Abfrage)

- Einspeisung als Stich aus dem HV: NHXHX-J 5×25 mm², NH00 gG 80 A, TN-C-S (PEN-Aufteilung im HV).
- Stockwerksverteiler Schrack Modul 2000 UP **IL166333WS** (3U-33G/WIEN, 6 Zählerplätze, 3 Nachzählerreihen, H1605×B810×T250).
- Je Zählerplatz: TYTAN-T-D02-Lasttrennschalter **IS503031** mit NEOZED 3×35 A gG (Allgemein 3×25 A), Zählersteckklemme **IL900189** (OVE R 21:2019), Smart Meter (Beistellung Wiener Netze), Hauptschalter 40 A 3-pol. **BZ900243**, Abgangsklemmen Phoenix TB 16 I.
- Allgemein: FI 40/4/0,03 Typ A,G, Stiegenhauslicht mit Treppenlichtzeitschalter VOWA+, Sprechanlage, Reserve. RWA bewusst ohne FI (vor dem FI abgegriffen).
- Wohnungsverteiler BK0850042 + BK0851042F (4×14 TE): Ableiter Typ 2 (+T3) **IS111314**, 2 FI-Gruppen **AR634103-A**, FI/LS **BO218616** für das Kühlgerät, LS AMPARO 10 kA, 18 Endstromkreise + 2 Reserve.

## Ergebnisse Auslegung

| Nachweis | Wert | Ergebnis |
|---|---|---|
| Leistungsbedarf / Ib | 47,8 kW / 72,6 A | ≤ 80 A |
| Leiterschutz W0 (Verlegeart C) | Iz 96 A, I2 128 A ≤ 139 A | erfüllt; bei Verlegeart B2/Häufung 5×35 mm² |
| ΔU W0 (30 m) / W1–W4 (25 m) | 0,83 % / 0,82 % | Grenzwert Hauptleitung lt. TAEV/Wiener Netze prüfen |
| Selektivität NH00 80 A → D02 35 A | 2,29 : 1 | selektiv |
| Ik am WV (10 kA am SV) | ≈ 3,4 kA | ≤ 6 kA FI |
| Verlustleistung SV / WV (RDF) | 79 W / 28 W | ≤ 455 W / 54 W |

## Vor Ausführung klären

- Ik am Einbauort, Grenzwert Spannungsfall Hauptleitung und Plombier-/Zählerplatzvorgaben mit Wiener Netze.
- Leitungslängen (Annahmen gelb in Excel) gegen den Installationsplan tauschen.
- Artikel mit Status „Schrack-Katalog“ (IS504037, IS504036, IS504136, IS504135, IKB04035, SI022270) vor Bestellung bestätigen. Meterware-Nummern bei Schrack anfragen.
- Brandschutzanforderung für den Verteiler im Stiegenhaus (OIB-RL 2) mit dem Brandschutzplaner abstimmen.

Ausführung nur durch ein konzessioniertes Elektrounternehmen (OVE E 8101).
