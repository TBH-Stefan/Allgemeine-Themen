# Huawei SUN2000-MAP0 + 2× LUNA2000-S0 + SmartGuard-63A-T0 (Whole-Home-Backup)

Anschlussplan inkl. Geräteeinstellungen für einen dreiphasigen Hybrid-WR SUN2000-MAP0 (FW V200R024C00SPC110) mit zwei Batteriestacks LUNA2000-5KW-C0 (V100R002C00SPC642) und SmartGuard-63A-T0 (V100R024C00SPC104).

Der SmartGuard übernimmt die Netz-/Inselumschaltung (≤ 20 ms mit MAP0), das Energiemanagement (integrierter EMMA) und die Messung am Netzanschluss. Deshalb kommt am MAP0 kein Smart Dongle und kein Smart Power Sensor zum Einsatz. Die Batterien hängen DC- und signalseitig am MAP0 und sind untereinander kaskadiert.

| Datei | Inhalt |
|---|---|
| `Anschlussplan_Huawei_MAP0_2xLUNA-S0_SmartGuard-T0.pdf` | Blatt 1: Übersichtsschaltplan · Blatt 2: Kabelliste, Schutzorgane, Absicherung je Leistungsklasse, Signalbelegung · Blatt 3: Parameter (FusionSolar), IBN-Reihenfolge, typische Fehler · Blatt 4: Erdungskonzept TN-C-S/TT, offene Fragen, Annahmen, Quellen |
| `src/` | `anschlussplan.html` → PDF über `node src/render_pdf.js` |

**Stand Index B (24.11.2026):** TN-C-S, Vorsicherung 35 A, gesamtes Haus am BACKUP-Port (NON-BACKUP frei), 2 Stacks à 3 × LUNA2000-5-E0 (30 kWh), keine Vorgaben des Netzbetreibers, kein Aggregat, kein Zeittarif, LAN eingeplant. Vorhandener Smart Dongle und Smart Power Sensor werden zurückgebaut.

**Index C (24.11.2026) – Steuerleitungen korrigiert:** SmartGuard-COM 1/2/3/4 ↔ MAP0-COM Pin 1/3/7/9 (RS485 + Freigabe; bisher fehlte 3/4 ↔ 7/9). Die Batteriekaskade LUNA1 → LUNA2 läuft 1:1 inklusive CANL/CANH (Pin 8/9). Am LUNA-Stecker gilt: um 180° gedreht, Wechselrichter an der rechten Seite. Abgeglichen mit der Wattkraft-Anleitung und laufenden Anlagen (Photovoltaikforum). Dazu eine Fehlersuch-Tabelle (Blatt 3, Abschn. 7).

**Index D:** MAP0 an SmartGuard-Port INV2 (32 A) lt. Huawei-Portzuordnungstabelle, INV1 bleibt frei.

**Noch offen:** Daten des vorhandenen FI vor dem HV (IΔn ≥ 100 mA, Typ A nötig – 30 mA muss versetzt werden), bestehende FusionSolar-Anlage, Großverbraucher für die Überlastbewertung im Inselbetrieb, Abstände Batterie–WR ≤ 10 m.

**Wichtig:** „Seamless switchover“ ist ab Werk deaktiviert. Die N-Leiter von Netz, WR, BACKUP und NON-BACKUP dürfen außerhalb des SmartGuard nicht verbunden sein. Pin-Belegungen und Parameterbezeichnungen stammen aus Huawei-Handbüchern (Ausgaben 2024–2026) und sind vor der Ausführung mit der Handbuch-Ausgabe des Geräts abzugleichen. Die nahtlose Umschaltung mit LUNA2000-S0 bei Huawei bestätigen lassen.

Ausführung nur durch eine konzessionierte Elektrofachkraft (OVE E 8101, TOR Erzeuger Typ A).
