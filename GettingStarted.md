# GettingStarted

Diese Anleitung hilft Ihnen Schritt für Schritt beim Einstieg in das öffentliche Repository.

## 1. Zugriff anfragen

Bevor Sie mit dem Repository arbeiten, fordern Sie bitte Zugriff an.

Senden Sie dazu eine E-Mail an:

**repo-access@example.org**

> Hinweis: Diese Adresse ist aktuell ein Platzhalter und kann später durch die echte Kontaktadresse ersetzt werden.

## 2. Repository öffnen oder klonen

Sobald Sie Zugriff haben, öffnen oder klonen Sie das Repository lokal.

Beispiel:

```bash
git clone <REPOSITORY-URL>
cd codat3_Validate
```

## 3. Repository-Struktur verstehen

Die wichtigsten Bereiche sind:

- `README.md`  
  Einstieg und Überblick
- `GettingStarted.md`  
  diese Onboarding-Anleitung
- `docs/validierungslogik.md`  
  fachliche Erklärung der Validierungslogik
- `docs/validierung-mit-github-actions.md`  
  Erklärung des GitHub-Validierungsablaufs
- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`  
  kanonische leere Vorlage
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`  
  kanonische leere Public-Vorlage
- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`  
  ausgefülltes allgemeines Beispiel
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`  
  ausgefülltes Public-Beispiel
- `scripts/validator/run_github_validation.py`  
  GitHub-kompatibler Einstiegspunkt für die Validierung
- `scripts/validator/validate_strukturvorlage.py`  
  zentrale Validierungslogik

Wichtig: Diese vier `.xlsx`-Dateien bilden gemeinsam die kanonische Vorlagenfamilie und müssen bei jeder fachlichen oder strukturellen Anpassung synchron nachgeführt werden.

## 4. Leere Vorlage herunterladen

Verwenden Sie für neue Arbeiten je nach Zielgruppe eine der beiden leeren Startvorlagen:

- `templates/Strukturvorlage_DataDictionary_empty_v0.9.4.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.4.xlsx`

Die Public-Variante enthält zusätzlich den Tab `Dictionary_public`.

## 5. Beispiel-Datei anschauen

Wenn Sie zuerst verstehen möchten, wie eine ausgefüllte Datei aussieht, öffnen Sie je nach Bedarf eines der Beispiele:

- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`
- `templates/test_files/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

Auch diese beiden Beispiel-Dateien müssen mit den leeren Vorlagen synchron bleiben.

## 6. Vorlage ausfüllen

Füllen Sie Ihre eigene Arbeitsmappe auf Basis der leeren Vorlage aus.

Wichtig:

- Blattnamen nicht umbenennen
- Kopfzeilen nicht verschieben
- Strukturblöcke nicht löschen
- Pflichtfelder im `Header` ausfüllen
- für Objektzuordnungen die Spalte `Classes.Class-Assignment` verwenden
- für Property-Zuordnungen die Spalte `Properties.Property-Assignment` verwenden
- die erlaubten Werte dafür nur aus den zugehörigen `Rules`-Spalten übernehmen:
  - `Rules.Class-Assignment`
  - `Rules.Property-Assignment`
- die zugehörigen Spalten
  - `Classes.Klassen-Zuordnung (DE)`
  - `Classes.Affectation de classe (FR)`
  - `Classes.Assegnazione della classe (IT)`
  - `Properties.Merkmals-Zuordnung (DE)`
  - `Properties.Attribution de propriété (FR)`
  - `Properties.Assegnazione della proprietà (IT)`
  sind nicht manuell zu pflegen, sondern system-generierte Ableitungen aus den `Rules`-Übersetzungsspalten

## 7. Validierung ausführen

Der öffentliche Standardweg ist die GitHub Action.

Grundablauf:

1. Ihre ausgefüllte `.xlsx` in einen Branch hochladen oder committen
2. GitHub Action **Validate Data Dictionary** starten
3. bei Bedarf `workbook_path` auf Ihre Datei setzen
4. Bericht und validierte Artefakt-`.xlsx` herunterladen und prüfen

## 8. Berichte lesen

Der Validator unterscheidet zwischen:

- **Fehler** = müssen behoben werden
- **Warnungen** = sollen geprüft werden
- **Normalisierungen** = zeigen automatische Ableitungen oder Standardisierungen

## 9. Relevante Dokumentation lesen

Für den Alltag sind diese Dateien besonders wichtig:

- `README.md`
- `docs/validierungslogik.md`
- `docs/validierung-mit-github-actions.md`

## 10. Mit dem Beispiel vergleichen

Wenn Ihre Datei nicht wie erwartet validiert, vergleichen Sie sie mit:

- der leeren Vorlage
- dem ausgefüllten AreaMgmt-Beispiel

So können Strukturfehler meist schnell erkannt werden.
