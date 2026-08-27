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

Die Struktur orientiert sich insbesondere an ISO 23386, ISO 23387, ISO 12006 und DCAT.

## Was Sie in diesem Repository finden

- eine **leere Startvorlage**
- ein **ausgefülltes Beispiel**
- eine **Validierungs-Pipeline**
- eine **Schritt-für-Schritt-Anleitung**
- ergänzende **deutschsprachige Dokumentation**
- maschinenlesbare **Validierungsberichte und validierte Arbeitsmappen** als GitHub-Artefakte

## Womit Sie starten sollen

Lesen Sie zuerst [GettingStarted.md](GettingStarted.md).

Verwenden Sie dann je nach Zielgruppe die passende leere Vorlage:

- [Allgemeine Vorlage](templates/Strukturvorlage_DataDictionary_empty_v1.0.0.xlsx)
- [Vorlage für öffentliche Auftraggeber](templates/Strukturvorlage_DataDictionary_empty_public_v0.9.5.xlsx)

Wenn Sie ein Beispiel brauchen, schauen Sie hier:

- [Area Management](templates/test_files/Strukturvorlage_AreaMgmt_v0.6.0.xlsx)
- [KBOB Facility Management](templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx)

## Validierungs-Artefakte

Wenn die Datei `pipeline_valid` ist, erzeugt die GitHub-Validierung zusätzlich eine validierte `.xlsx`-Artefaktdatei mit system-generierten Rückschreibungen, zum Beispiel für abgeleitete IDs und andere sichere Normalisierungen. Diese Datei wird zusammen mit den JSON- und Markdown-Berichten als GitHub-Artefakt hochgeladen.

## Nutzen

Wenn Sie die Validierung konsequent durchlaufen, erhalten Sie:

- bessere Datenqualität
- klarere Begriffe und Definitionen
- konsistentere Struktur
- sauberere Referenzen zwischen Klassen, Merkmalen, Werten und Dokumenten
- eine bessere Grundlage für digitale Weiterverwendung und RDF-Publikation
