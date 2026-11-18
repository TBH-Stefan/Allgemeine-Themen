# PV-Überschussnutzung SolarEdge SE8K → WOLF CHA-10 (BM-2, FW 3.10)

WW-Soll bleibt bei 50 °C. Bei PV-Überschuss wird der WW-Speicher zweistufig weiter beheizt:

1. **Wärmepumpe:** Der SolarEdge Home Load Controller (−A1) schaltet über das Koppelrelais −K1 den WOLF-Eingang X0 PV/SG1–GND. Damit hebt die CHA das WW-Soll um WP027 = 10 K auf 60 °C an.
2. **E-Heizpatrone:** Der SolarEdge Home Hot Water Controller (−A2) gibt den restlichen Überschuss stufenlos (0–3 kW) an die Heizpatrone, bis deren Thermostat bei 75 °C abschaltet.

Über die WOLF-Regelung allein sind 75 °C nicht erreichbar: WP027 erlaubt max. +20 K, A14 begrenzt zusätzlich, und das E-Heizelement sitzt im Heizwasserkreis.

| Datei | Inhalt |
|---|---|
| `Anschlussplan_PV-Ueberschuss_SE8K_WOLF-CHA10.pdf` | Blatt 1: Übersichtsschaltplan · Blatt 2: Funktion, Klemmen- und Kabelliste · Blatt 3: Parameter WOLF/SolarEdge, IBN-Checkliste, Annahmen |
| `Auslegung_Stueckliste_PV-Ueberschuss.xlsx` | Auslegung (Energie, Schwellen, Schutzorgan, ΔU, Wirtschaftlichkeit, Sensitivität), Stückliste mit Optionen, Klemmenplan, Parameter, IBN |
| `src/` | Quellen: `anschlussplan.html` → PDF über `node src/render_pdf.js`; `python3 src/build_xlsx.py` erzeugt die Excel-Datei |

**Vor Ausführung klären:** Daten der Heizpatrone (Phasen, Leistung, TR-Bereich), zulässige Speichertemperatur, Wasserhärte (> 15 °dH → Temperatur begrenzen), ob bereits ein SolarEdge Energy Meter vorhanden ist, Softwarestand von HCM-4.

Ausführung nur durch eine konzessionierte Elektrofachkraft (OVE E 8101). Die WOLF-Parameternummern bei BM-2 FW 3.10 vor Ort verifizieren.
