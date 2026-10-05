# Repo-Nutzung und Navigation

Diese Datei erklärt, wie sich Benutzer im öffentlichen Repository schnell orientieren können.

## Ziel des Repositorys

Das Repository stellt eine **öffentliche Validierungs-Pipeline** für Data Dictionary-Arbeitsmappen bereit.

Es soll Benutzern helfen,

- die offizielle Vorlage zu verwenden,
- ein Beispiel zu verstehen,
- die Validierung korrekt auszuführen,
- und die Berichte richtig zu lesen.

## Empfohlene Lesereihenfolge

1. `GettingStarted.md`
2. `docs/strukturvorlage-schritt-fuer-schritt.md`
3. `docs/repository-housekeeping.md`
4. `README.md`
5. `docs/validierungslogik.md`
6. `docs/validierung-mit-github-actions.md`

## Wichtigste Dateien

### Für neue Nutzer

- `GettingStarted.md`
- `README.md`
- `docs/strukturvorlage-schritt-fuer-schritt.md`
- `docs/repository-housekeeping.md`
- `docs/Leitfaden_Strukturvorlage_Housekeeping_und_Anleitung.pdf`

### Für die Arbeit mit Excel-Dateien

- `templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`
- `WIP data dictionaries/IFMA/CHE.5539 IFMA_Data Template_AreaMgmt_v0.9.5.xlsx`
- `WIP data dictionaries/KBOB/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

Ausgefüllte Data Dictionaries durchlaufen die Statusbereiche `WIP data dictionaries`, `SHARED data dictionaries`, `PUBLISHED data dictionaries` und `ARCHIVED data dictionaries`. Innerhalb jedes Bereichs erfolgt die Ablage nach Organisation.

### Für das Verständnis der Validierung

- `docs/validierungslogik.md`
- `scripts/validator/validate_strukturvorlage.py`

### Für die GitHub-Ausführung

- `.github/workflows/validate-data-dictionary.yml`
- `scripts/validator/run_github_validation.py`
- `docs/validierung-mit-github-actions.md`
- `validator_tests/` für automatisierte Validator-Regressionstests

## Was Benutzer im Normalfall nicht anfassen müssen

Normale Nutzer müssen in der Regel nicht direkt arbeiten mit:

- internen Python-Hilfsmodulen,
- Reader-/Writer-Unterstrukturen,
- Archivmaterial,
- lokalen Entwicklungsdateien.

## Grundprinzip

Für die öffentliche Nutzung gilt:

- Vorlage nehmen
- Datei ausfüllen
- Validierung starten
- Bericht lesen
- Fehler korrigieren
