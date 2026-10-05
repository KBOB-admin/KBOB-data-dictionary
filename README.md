# Öffentliche Data-Dictionary-Validierung

> **Wichtiger Hinweis:** Das Repository `KBOB-data-dictionary` ist ein Arbeitsrepository. Es werden **keine Supportleistungen**, **keine Hilfestellung**, **keine Kontaktadresse** und **kein betreuter Kommunikationskanal** angeboten. Die Nutzung der bereitgestellten Dateien und Hinweise geschieht auf eigene Verantwortung. 

Dieses Repository stellt Excel-Vorlagen und eine automatisierte Validierung bereit, um bestehende Data Dictionaries aus Planung, Bau und Betrieb von Bauwerken strukturiert zu erfassen und ihre Qualität zu verbessern.

## Zusammenspiel der KBOB-Repositories

- [`KBOB-data-dictionary`](https://github.com/KBOB-admin/KBOB-data-dictionary) pflegt Vorlagen und Validierungslogik.
- [`KBOB-data-dictionary-schemaforge`](https://github.com/KBOB-admin/KBOB-data-dictionary-schemaforge) transformiert validierte Arbeitsmappen in RDF- und I14Y-Publikationsartefakte.
- [`KBOB-data-dictionary-schema`](https://github.com/KBOB-admin/KBOB-data-dictionary-schema) ist die normative Quelle für das von SchemaForge verwendete NatDD-`dd:`-Kernvokabular.

Die Vorlagen legen die fachliche Eingabestruktur fest. Der Validator prüft diese Struktur, definiert aber keine RDF-Begriffe. Die Bedeutung der publizierten `dd:`-Klassen und -Properties wird ausschliesslich im Schema-Repository gepflegt. Dessen Namespace ist während der öffentlichen `0.x`-Reviewphase noch provisorisch.

## Vorgehen

1. Leere Vorlage herunterladen.
2. Inhalte des bestehenden Data Dictionary eintragen.
3. Datei über GitHub Actions validieren.
4. Fehler und Warnungen anhand des Validierungsberichts bearbeiten.
5. Validierte Arbeitsmappe für die weitere Publikation verwenden.

Das Ergebnis ist ein klar strukturiertes und konsistentes Data Dictionary mit nachvollziehbaren Bezeichnungen, Definitionen und Referenzen.

Die Struktur orientiert sich insbesondere an ISO 23386, ISO 23387, ISO 12006 und DCAT. Die Ablage der Data Dictionaries folgt zusätzlich einem an ISO 19650 angelehnten Statusmodell:

- [`WIP data dictionaries`](WIP%20data%20dictionaries/) – in Bearbeitung, noch nicht freigegeben
- [`SHARED data dictionaries`](SHARED%20data%20dictionaries/) – zur Koordination und Prüfung freigegeben
- [`PUBLISHED data dictionaries`](PUBLISHED%20data%20dictionaries/) – für die Nutzung autorisiert
- [`ARCHIVED data dictionaries`](ARCHIVED%20data%20dictionaries/) – abgelöst oder zurückgezogen; Audit-Trail

Innerhalb jedes Statusbereichs werden die Dateien nach verantwortlicher Organisation abgelegt. Leere Ausgangsvorlagen verbleiben unabhängig vom Statusmodell unter [`templates/`](templates/).

## Was Sie in diesem Repository finden

- eine **leere Startvorlage**
- **Data Dictionaries nach Bearbeitungsstatus und Organisation**
- eine **Validierungs-Pipeline**
- eine [**Schritt-für-Schritt-Anleitung**](docs/strukturvorlage-schritt-fuer-schritt.md)
- ergänzende **deutschsprachige Dokumentation**
- maschinenlesbare **Validierungsberichte und validierte Arbeitsmappen** als GitHub-Artefakte
- verbindliche [Housekeeping- und Ablageregeln](docs/repository-housekeeping.md)
- einen zusammengeführten [PDF-Leitfaden](docs/Leitfaden_Strukturvorlage_Housekeeping_und_Anleitung.pdf)

Der Ordner [`validator_tests/`](validator_tests/) enthält ausschliesslich
automatisierte Regressionstests für den Validator. Data-Dictionary-Dateien
dürfen dort nicht abgelegt werden.

## Womit Sie starten sollen

Lesen Sie zuerst die
[Schritt-für-Schritt-Anleitung](docs/strukturvorlage-schritt-fuer-schritt.md)
und die [Housekeeping-Regeln](docs/repository-housekeeping.md).

Verwenden Sie die aktuelle leere Vorlage:

- [Strukturvorlage Data Dictionary v1.1.0](templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx)

Wenn Sie ein ausgefülltes Data Dictionary brauchen, schauen Sie im WIP-Bereich der jeweiligen Organisation:

- [IFMA – Area Management v0.9.5](WIP%20data%20dictionaries/IFMA/CHE.5539%20IFMA_Data%20Template_AreaMgmt_v0.9.5.xlsx)
- [KBOB – Facility Management](WIP%20data%20dictionaries/KBOB/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx)

## Validierungs-Artefakte

Wenn die Datei `pipeline_valid` ist, erzeugt die GitHub-Validierung zusätzlich eine validierte `.xlsx`-Artefaktdatei mit system-generierten Rückschreibungen, zum Beispiel für abgeleitete IDs und andere sichere Normalisierungen. Diese Datei wird zusammen mit den JSON- und Markdown-Berichten als GitHub-Artefakt hochgeladen.

## Nutzen

Wenn Sie die Validierung konsequent durchlaufen, erhalten Sie:

- bessere Datenqualität
- klarere Begriffe und Definitionen
- konsistentere Struktur
- sauberere Referenzen zwischen Klassen, Merkmalen, Werten und Dokumenten
- eine bessere Grundlage für digitale Weiterverwendung und RDF-Publikation
