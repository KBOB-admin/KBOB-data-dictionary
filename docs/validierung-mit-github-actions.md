# Validierung mit GitHub Actions

## Standardpfad der GitHub Action

Die GitHub Action verwendet standardmässig:

- `templates/test_files/Strukturvorlage_AreaMgmt_v0.5.0.xlsx`

Dieser Standardpfad dient als öffentlicher Referenztest für eine ausgefüllte und pipeline-gültige Beispielarbeitsmappe.

## Empfohlener Nutzer-Workflow

1. Vorlage herunterladen
2. Vorlage ausfüllen
3. Ausgefüllte `.xlsx` in einen Branch hochladen oder committen
4. Workflow **Validate Data Dictionary** manuell starten
5. Optional `workbook_path` auf den relativen Pfad der eigenen Datei setzen
6. Validierungsbericht und Artefakte herunterladen und lesen

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

## Aktueller Geltungsbereich

Der öffentliche MVP deckt die **Validierung von ausgefüllten Arbeitsmappen** ab.

Nicht Teil dieses öffentlichen MVP sind:

- RDF-Export
- bSDD-Publikation
- i14y-Publikation
- LINDAS-Publikation
- sonstige Export-/Publishing-Pipelines
