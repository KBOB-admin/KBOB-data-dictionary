# Öffentliche Data-Dictionary-Validierung

> **Wichtiger Hinweis:** Das Repository `KBOB-data-dictionary` ist derzeit ein Arbeitsrepository. Aktuell werden **keine Supportleistungen**, **keine individuelle Hilfestellung**, **keine Kontaktadresse** und **kein betreuter Kommunikationskanal** angeboten. Bitte nutzen Sie die bereitgestellten Dateien und Hinweise ausschliesslich in eigener Verantwortung.

Dieses Repository hilft Organisationen dabei, ihren bestehenden Datenkatalog in eine strukturierte Excel-Vorlage zu überführen und die Qualität dieser Daten systematisch zu verbessern.

Kurz gesagt ist der Ablauf so:

1. die leere Vorlage herunterladen
2. den heutigen Datenkatalog des Unternehmens eintragen
3. die Datei über GitHub Actions validieren
4. Fehler und Warnungen schrittweise bereinigen
5. so lange verbessern, bis die Datei sauber durchläuft

Das Ziel ist nicht einfach nur eine "gültige Datei", sondern ein besser strukturierter, klarer und konsistenter Datenkatalog.

Dabei hilft die Validierung, den Inhalt schrittweise an bewährte fachliche und semantische Standards anzunähern, insbesondere an:

- ISO 23386 / ISO 23387
- ISO 12006
- DCAT

## Was Sie in diesem Repository finden

- eine **leere Startvorlage**
- ein **ausgefülltes Beispiel**
- eine **Validierungs-Pipeline**
- eine **Schritt-für-Schritt-Anleitung**
- ergänzende **deutschsprachige Dokumentation**

## Womit Sie starten sollen

Lesen Sie zuerst:

- `GettingStarted.md`

Verwenden Sie dann je nach Zielgruppe die passende leere Vorlage:

- `templates/Strukturvorlage_DataDictionary_empty.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public.xlsx`

Wenn Sie ein Beispiel brauchen, schauen Sie hier:

- `templates/test_files/Data Dictionary_BdCH_AreaMgmt.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM.xlsx`

## Struktur der Vorlagen

Es gibt zwei leere Hauptvorlagen:

- `*_empty.xlsx`
- `*_empty_public.xlsx`

### `*_empty.xlsx`

Diese Vorlage ist für den allgemeinen industriellen Einsatz gedacht.
Sie eignet sich für Unternehmen und Organisationen, die einen Data Dictionary intern oder branchenübergreifend strukturieren wollen, ohne den zusätzlichen öffentlichen Metadatenblock für öffentliche Auftraggeber zu benötigen.

### `*_empty_public.xlsx`

Diese Vorlage ist für öffentliche Auftraggeber und öffentliche Kunden gedacht.
Sie enthält zusätzlich den Tab `Dictionary_public` und unterstützt damit weitergehende öffentliche Metadaten für Veröffentlichung, Governance und Katalogisierung.

### Beispielvorlagen

Zusätzlich gibt es zwei ausgefüllte Beispielvorlagen:

- `Data Dictionary_BdCH_AreaMgmt.xlsx` als allgemeines Beispiel
- `Strukturvorlage_DataDictionary_KBOB_FM.xlsx` als Beispiel für öffentliche Auftraggeber

## Wichtiger Nutzen

Wenn Sie die Validierung konsequent durchlaufen, erhalten Sie:

- bessere Datenqualität
- klarere Begriffe und Definitionen
- konsistentere Struktur
- sauberere Referenzen zwischen Klassen, Merkmalen, Werten und Dokumenten
- eine bessere Grundlage für spätere digitale Weiterverwendung

## Öffentlicher Nutzungszweck

Dieses Repository ist für eine saubere öffentliche Nutzung gedacht.

Es enthält deshalb nur:

- die öffentliche Vorlage
- das öffentliche Beispiel
- die nutzbare Validierungslogik
- unterstützende Dokumentation
