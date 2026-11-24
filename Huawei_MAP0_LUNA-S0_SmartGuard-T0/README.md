# Huawei SUN2000-MAP0 + 2× LUNA2000-S0 + SmartGuard-63A-T0 (Whole-Home-Backup)

Anschlussplan inkl. Geräteeinstellungen für einen dreiphasigen Hybrid-WR SUN2000-MAP0 (FW V200R024C00SPC110) mit zwei Batteriestacks LUNA2000-5KW-C0 (V100R002C00SPC642) und SmartGuard-63A-T0 (V100R024C00SPC104).

Der SmartGuard übernimmt die Netz-/Inselumschaltung (≤ 20 ms mit MAP0), das Energiemanagement (integrierter EMMA) und die Messung am Netzanschluss. Deshalb kommt am MAP0 kein Smart Dongle und kein Smart Power Sensor zum Einsatz. Die Batterien hängen DC- und signalseitig am MAP0 und sind untereinander kaskadiert.

| Datei | Inhalt |
|---|---|
| `Anschlussplan_Huawei_MAP0_2xLUNA-S0_SmartGuard-T0.pdf` | Blatt 1: Übersichtsschaltplan · Blatt 2: Kabelliste, Schutzorgane, Absicherung je Leistungsklasse, Signalbelegung · Blatt 3: Parameter (FusionSolar), IBN-Reihenfolge, typische Fehler · Blatt 4: Erdungskonzept TN-C-S/TT, offene Fragen, Annahmen, Quellen |
| `src/` | `anschlussplan.html` → PDF über `node src/render_pdf.js` |

**Vor Ausführung klären:** Leistungsklasse MAP0, Module je Stack, Netzform TN-C-S/TT und Ort der PEN-Aufteilung, Vorsicherung (≤ 63 A für Whole-Home), Aufteilung BACKUP/NON-BACKUP, Vorgaben des Netzbetreibers (Einspeiselimit, Blindleistung), Internetanbindung.

**Wichtig:** „Seamless switchover“ ist ab Werk deaktiviert. Die N-Leiter von Netz, WR, BACKUP und NON-BACKUP dürfen außerhalb des SmartGuard nicht verbunden sein. Pin-Belegungen und Parameterbezeichnungen stammen aus Huawei-Handbüchern (Ausgaben 2024–2026) und sind vor der Ausführung mit der Handbuch-Ausgabe des Geräts abzugleichen. Die nahtlose Umschaltung mit LUNA2000-S0 bei Huawei bestätigen lassen.

Ausführung nur durch eine konzessionierte Elektrofachkraft (OVE E 8101, TOR Erzeuger Typ A).
