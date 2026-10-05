# GettingStarted

Diese Anleitung hilft Ihnen Schritt für Schritt beim Einstieg in das öffentliche Repository.

## 1. Beitragsweg wählen

Für Beiträge ist keine Mitgliedschaft und kein direkter Schreibzugriff auf das
Repository erforderlich. Der Standardweg ist ein eigener Fork mit
anschliessendem Pull Request.

Falls eine Organisation diesen GitHub-Weg nicht selbst durchführen kann, darf
sie als Fallback eine Anfrage an folgende Adresse senden:

**info@intop3.ch**

Dieses Postfach wird einmal pro Woche bearbeitet. Es kann deshalb zu einer
Wartezeit von bis zu sieben Kalendertagen kommen. Die E-Mail-Anfrage begründet
weder einen Anspruch auf Aufnahme noch eine fachliche Freigabe der
eingereichten Inhalte.

## 2. Repository forken oder klonen

Erstellen Sie für einen Beitrag zunächst einen Fork des Repositorys. Arbeiten
Sie ausschliesslich in diesem Fork. Ein direkter Schreibzugriff auf das
Original-Repository wird nicht benötigt.

Beispiel:

```bash
git clone <REPOSITORY-URL>
cd KBOB-data-dictionary
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
- `docs/repository-housekeeping.md`
  verbindliche Regeln für Ablage, Statuswechsel und Verantwortlichkeiten
- `docs/strukturvorlage-schritt-fuer-schritt.md`
  vollständiger Ablauf von der Vorlage bis zum Pull Request
- `templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`
  aktuelle leere Strukturvorlage
- `WIP data dictionaries/IFMA/CHE.5539 IFMA_Data Template_AreaMgmt_v0.9.5.xlsx`
  aktuelles IFMA Data Dictionary in Bearbeitung
- `WIP data dictionaries/KBOB/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`
  ausgefülltes KBOB Data Dictionary in Bearbeitung
- `scripts/validator/run_github_validation.py`  
  GitHub-kompatibler Einstiegspunkt für die Validierung
- `scripts/validator/validate_strukturvorlage.py`  
  zentrale Validierungslogik
- `validator_tests/`
  automatisierte Regressionstests des Validators; keine Ablage für Data Dictionaries

Die leeren Vorlagen verbleiben unter `templates/`. Ausgefüllte Data Dictionaries werden nach Bearbeitungsstatus (`WIP`, `SHARED`, `PUBLISHED`, `ARCHIVED`) und danach nach verantwortlicher Organisation abgelegt.

## 4. Leere Vorlage herunterladen

Verwenden Sie für neue Arbeiten die aktuelle leere Startvorlage:

- `templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`

## 5. Beispiel-Datei anschauen

Wenn Sie zuerst verstehen möchten, wie eine ausgefüllte Datei aussieht, öffnen Sie je nach Bedarf eines der Beispiele:

- `WIP data dictionaries/IFMA/CHE.5539 IFMA_Data Template_AreaMgmt_v0.9.5.xlsx`
- `WIP data dictionaries/KBOB/Strukturvorlage_DataDictionary_KBOB_FM_v0.9.5.xlsx`

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

System-generierte Felder werden erst dann in die validierte Artefaktdatei
geschrieben, wenn keine Blocking Errors mehr vorhanden sind.

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
