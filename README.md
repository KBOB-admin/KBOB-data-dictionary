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

Aktuelle Kernbegriffe der Vorlage sind unter anderem:

- `Classes.Class-Assignment`
- `Classes.Klassen-Zuordnung (DE)`
- `Classes.Affectation de classe (FR)`
- `Classes.Assegnazione della classe (IT)`
- `Properties.Property-Assignment`
- `Properties.Merkmals-Zuordnung (DE)`
- `Properties.Attribution de propriété (FR)`
- `Properties.Assegnazione della proprietà (IT)`
- `Rules.Class-Assignment`
- `Rules.Klassen-Zuordnung`
- `Rules.Affectation de classe`
- `Rules.Assegnazione della classe`
- `Rules.Property-Assignment`
- `Rules.Merkmals-Zuordnung`
- `Rules.Attribution de propriété`
- `Rules.Assegnazione della proprietà`

Die englischen Assignment-Spalten sind die führenden Eingabefelder. Die zugehörigen DE/FR/IT-Spalten in `Classes` und `Properties` sind system-generierte Ableitungen aus den entsprechenden `Rules`-Übersetzungsspalten.

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

- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`

Wenn Sie ein Beispiel brauchen, schauen Sie hier:

- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

## Struktur der Vorlagen

Es gibt vier kanonische `.xlsx`-Dateien, die fachlich und strukturell synchron gehalten werden müssen:

### Leere Hauptvorlagen

- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`

`*_empty.xlsx` ist die allgemeine Vorlage ohne öffentlichen Zusatzblock.

`*_empty_public.xlsx` ist die Variante für öffentliche Auftraggeber. Sie enthält zusätzlich den Tab `Dictionary_public`.

### Ausgefüllte Beispielvorlagen

- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

Diese beiden Beispiel-Dateien müssen dieselbe fachliche Struktur, dieselben Blattnamen, dieselben Kernspalten und dieselbe Guidance widerspiegeln wie die leeren Hauptvorlagen. Die einzige zulässige strukturelle Abweichung bleibt der öffentliche Zusatzblock `Dictionary_public` in den beiden Public-Varianten.

## Synchronisationsregel für künftiges Feedback

Wenn Feedback an der Vorlage umgesetzt wird, muss es immer auf alle vier kanonischen `.xlsx`-Dateien angewendet werden:

- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`
- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

Dabei gilt:

- die beiden leeren Vorlagen bleiben untereinander synchron,
- die beiden Beispielvorlagen bleiben untereinander synchron,
- Beispiele und leere Vorlagen bleiben in Blattstruktur, Kernspalten, Benennungen und Guidance synchron,
- nur der Public-Zusatz `Dictionary_public` darf exklusiv in den Public-Dateien bestehen.

## Validator-Ausrichtungsregel

Jede strukturelle oder fachliche Änderung an den Vorlagen muss gleichzeitig im Validator nachvollzogen werden.

Das bedeutet insbesondere:

- keine veralteten Blattnamen im Validator belassen,
- keine Legacy-Spaltennamen als primäre Zielstruktur weiterführen,
- keine Prüfregeln für Strukturen behalten, die in den synchronisierten Vorlagen nicht mehr existieren,
- Dokumentation, GitHub-Validierung und Validator-Logik gemeinsam nachführen.

Ziel ist, dass der Validator immer die aktuell synchronisierte Vorlagenfamilie prüft und keine historischen Artefakte künstlich konserviert.

## Wichtige Benennungsregel

Die aktuelle Vorlagenfamilie verwendet die folgenden Zuordnungsbegriffe:

- `Classes.Class-Assignment` statt `Classes.Classification`
- `Properties.Property-Assignment` statt `Properties.Property Classification`
- `Rules.Property-Assignment` statt `Rules.Property Classification`
- `Rules.Class-Assignment` als englische Übersetzungs- und Dropdown-Spalte zu `Rules.Klassifikation`

Wenn Sie bestehende Arbeitsmappen oder ältere Beispiele vergleichen, achten Sie darauf, nur die aktuelle Benennung der synchronisierten Vorlagenfamilie zu verwenden.

## Validierungs-Artefakte

Wenn die Datei `pipeline_valid` ist, erzeugt die GitHub-Validierung zusätzlich eine validierte `.xlsx`-Artefaktdatei mit system-generierten Rückschreibungen, zum Beispiel für abgeleitete IDs und andere sichere Normalisierungen. Diese Datei wird zusammen mit den JSON- und Markdown-Berichten als GitHub-Artefakt hochgeladen.

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
