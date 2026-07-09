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
2. `README.md`
3. `docs/validierungslogik.md`
4. `docs/validierung-mit-github-actions.md`

## Wichtigste Dateien

### Für neue Nutzer

- `GettingStarted.md`
- `README.md`

### Für die Arbeit mit Excel-Dateien

- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`
- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

### Für das Verständnis der Validierung

- `docs/validierungslogik.md`
- `scripts/validator/validate_strukturvorlage.py`

### Für die GitHub-Ausführung

- `.github/workflows/validate-data-dictionary.yml`
- `scripts/validator/run_github_validation.py`
- `docs/validierung-mit-github-actions.md`

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
