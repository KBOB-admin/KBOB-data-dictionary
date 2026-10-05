# Validierung mit GitHub Actions

## Standardpfad der GitHub Action

Die GitHub Action verwendet standardmässig:

- `templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`

Dieser Standardpfad dient als struktureller Referenztest für die aktuelle leere Vorlage. Fehlende, vom Benutzer einzutragende Metadaten werden bei dieser Datei als Warnungen gemeldet.

## Empfohlener Nutzer-Workflow

1. Vorlage herunterladen
2. Vorlage ausfüllen
3. Ausgefüllte `.xlsx` in einen Branch hochladen oder committen
4. Workflow **Validate Data Dictionary** manuell starten
5. Optional `workbook_path` auf den relativen Pfad der eigenen Datei setzen
6. Validierungsbericht und Artefakte herunterladen und lesen

Externe Beitragende arbeiten im eigenen Fork und reichen ihre Datei über einen
Pull Request ein. Sie benötigen keinen direkten Schreibzugriff auf das
Original-Repository. Der vollständige Ablauf ist in
`docs/strukturvorlage-schritt-fuer-schritt.md` beschrieben.

## Öffentliche Blattnamen

Die öffentliche Vorlage und die Validierung erwarten im MVP diese Blattnamen:

- `Header`
- `Classes`
- `Properties`
- `Values`
- `Documents`
- `GroupOfProperties`
- `Rules`
- `Data_Template`

## Wo die Berichte landen

Die Validierung erzeugt Berichte unter:

- `Validation_output/`

In GitHub Actions werden diese Berichte zusätzlich als Artefakte hochgeladen. Wenn die Arbeitsmappe `pipeline_valid` ist, wird ausserdem eine validierte `.xlsx` mit system-generierten Rückschreibungen in `Validation_output/` erzeugt und als Artefakt mit hochgeladen.

Solange mindestens ein Blocking Error besteht, wird keine validierte
Arbeitsmappe mit system-generierten Rückschreibungen erzeugt.

## Aktueller Geltungsbereich

Der öffentliche MVP deckt die **Validierung von ausgefüllten Arbeitsmappen** ab.

Nicht Teil dieses öffentlichen MVP sind:

- RDF-Export
- bSDD-Publikation
- i14y-Publikation
- LINDAS-Publikation
- sonstige Export-/Publishing-Pipelines
