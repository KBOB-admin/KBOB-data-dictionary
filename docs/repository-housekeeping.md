# Housekeeping- und Ablageregeln für Data Dictionaries

## 1. Zweck und Geltungsbereich

Dieses Repository stellt eine einheitliche Strukturvorlage, eine geordnete
Ablage und einen automatisierten Validierungsdienst für Data Dictionaries zur
Verfügung. Die Regeln gelten für alle eingereichten und im Repository
verwalteten Strukturvorlagen.

Der Validierungsdienst prüft die strukturelle und technisch prüfbare
Konsistenz einer Arbeitsmappe. Dazu gehören insbesondere Pflichtfelder,
Tabellen- und Spaltenstrukturen, Referenzen, kontrollierte Wertelisten,
Datentypen, Formate und weitere maschinenlesbare Regeln.

Die Validierung ist keine fachliche Freigabe und keine Bestätigung der
inhaltlichen Richtigkeit, Vollständigkeit, Normkonformität, rechtlichen
Verwendbarkeit oder Eignung für einen bestimmten Zweck.

## 2. Verbindliche Ausgangsvorlage

Für neue Data Dictionaries ist ausschliesslich die jeweils im Ordner
`templates/` als aktuell bezeichnete leere Strukturvorlage zu verwenden.
Derzeit ist dies:

`templates/2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`

Die Originaldatei unter `templates/` darf nicht mit organisations- oder
projektspezifischen Inhalten überschrieben werden. Für die Bearbeitung ist eine
Kopie zu erstellen. Blattnamen, Kopfzeilen, Strukturblöcke, technische
Referenzen und Validierungsbereiche dürfen nicht eigenmächtig verändert,
verschoben oder gelöscht werden.

Bestehende Data Dictionaries auf älteren Vorlagen sind vor einer neuen
Einreichung grundsätzlich in die aktuelle Strukturvorlage zu überführen. Eine
Ausnahme ist nur zulässig, wenn die ältere Version aus Gründen der
Nachvollziehbarkeit unverändert archiviert wird.

## 3. Verantwortlichkeiten

Der jeweilige Data-Dictionary-Owner trägt die alleinige Verantwortung für:

- die fachliche Richtigkeit und Vollständigkeit der Inhalte;
- Definitionen, Benennungen, Übersetzungen, Klassifikationen und Wertelisten;
- Rechte an den eingereichten Inhalten und Quellen;
- fachliche Abstimmung, Qualitätssicherung und Freigabe;
- Versionsentscheidungen und den vorgesehenen Verwendungszweck;
- die Entscheidung, ob ein Data Dictionary von `WIP` nach `SHARED` oder
  `PUBLISHED` überführt werden darf.

Die Repository-Verwaltung ist verantwortlich für:

- die technische Ordnung und nachvollziehbare Ablage im Repository;
- die Bereitstellung der aktuellen Strukturvorlage;
- die Ausführung und Pflege des automatisierten Validierungsworkflows;
- die technische Kontrolle von Ordner, Dateiname, Version und Änderungsumfang;
- den Schutz des Hauptbranches durch einen geregelten Review- und
  Freigabeprozess.

Die Repository-Verwaltung übernimmt durch Annahme einer Datei keine
Verantwortung für deren fachlichen Inhalt. Ein bestandener Validierungslauf
ersetzt weder die fachliche Prüfung noch die ausdrückliche Freigabe durch den
Data-Dictionary-Owner.

## 4. Statusbereiche und zulässige Ablage

Jede ausgefüllte Strukturvorlage wird innerhalb des zutreffenden Statusbereichs
in einem Unterordner der verantwortlichen Organisation abgelegt:

`<STATUS> data dictionaries/<Organisation>/<Dateiname>.xlsx`

### 4.1 WIP data dictionaries

`WIP` enthält Data Dictionaries in aktiver Bearbeitung.

- Die Inhalte sind nicht freigegeben.
- Blocking Errors und Warnungen dürfen noch vorhanden sein.
- Die Datei darf nicht als geprüft, publiziert oder für den operativen Einsatz
  geeignet bezeichnet werden.
- Jede Organisation verwendet einen eindeutig bezeichneten eigenen
  Unterordner.
- Zwischenstände dürfen aufbewahrt werden, wenn Dateiname und Version eindeutig
  unterscheidbar sind.

### 4.2 SHARED data dictionaries

`SHARED` enthält Data Dictionaries, die der Data-Dictionary-Owner ausdrücklich
zur Koordination oder fachlichen Prüfung freigegeben hat.

Voraussetzungen für den Statuswechsel von `WIP` nach `SHARED`:

- keine Blocking Errors in der automatisierten Validierung;
- Prüfung und bewusste Behandlung aller Warnungen;
- vollständige Owner-, Versions- und Provenance-Angaben;
- dokumentierte Freigabe zur Koordination durch den Data-Dictionary-Owner.

`SHARED` bedeutet nicht, dass das Data Dictionary für den operativen Einsatz
autorisiert ist.

### 4.3 PUBLISHED data dictionaries

`PUBLISHED` enthält formal freigegebene Versionen.

Voraussetzungen für den Statuswechsel nach `PUBLISHED`:

- alle Anforderungen des Status `SHARED` sind erfüllt;
- die fachliche Qualitätssicherung ist abgeschlossen;
- der Data-Dictionary-Owner hat die konkrete Version ausdrücklich zur Nutzung
  freigegeben;
- Versionsnummer, Status, Datum, Verantwortlichkeit und Provenance sind
  nachvollziehbar dokumentiert;
- die publizierte Datei entspricht genau der freigegebenen und validierten
  Version.

Eine Datei darf nicht allein aufgrund eines technisch erfolgreichen
Validierungslaufs nach `PUBLISHED` verschoben werden.

### 4.4 ARCHIVED data dictionaries

`ARCHIVED` enthält abgelöste, zurückgezogene oder nur noch zu
Nachweiszwecken aufbewahrte Versionen.

- Archivierte Dateien dürfen nicht als aktuelle Version verwendet werden.
- Dateien werden nicht gelöscht, wenn sie für den Audit-Trail oder die
  Nachvollziehbarkeit einer früheren Freigabe erforderlich sind.
- Der Grund der Archivierung soll über Commit, Pull Request oder begleitende
  Dokumentation nachvollziehbar sein.

## 5. Dateinamen und Versionierung

Dateinamen müssen Organisation, Data Dictionary und Version eindeutig erkennen
lassen. Empfohlenes Schema:

`<Organisation>_<Data-Dictionary-Name>_v<MAJOR.MINOR.PATCH>.xlsx`

Für eine fachlich geänderte Datei ist eine neue Version zu vergeben. Eine
bereits publizierte Version darf nicht stillschweigend durch anderen Inhalt
ersetzt werden. Temporäre Office-Dateien, insbesondere Dateien mit dem Präfix
`~$`, dürfen nicht committed werden.

Eine Datei soll grundsätzlich nur in einem aktiven Statusbereich geführt
werden. Bei einem Statuswechsel ist sie zu verschieben, nicht als unabhängige
Kopie in mehreren Statusbereichen weiterzuführen. Abgelöste freigegebene
Versionen gehören nach `ARCHIVED`.

## 6. Validierung und Quality Gates

Vor jedem Statuswechsel und vor jeder formalen Freigabe ist die GitHub Action
`Validate Data Dictionary` für die konkrete Arbeitsmappe auszuführen.

Die Ergebnisse sind wie folgt zu behandeln:

- **Blocking Error:** muss vor `SHARED` oder `PUBLISHED` behoben sein;
- **Warnung:** muss durch den Data-Dictionary-Owner geprüft und bewusst
  behandelt werden;
- **Normalisierung:** zeigt eine sichere technische Ableitung oder
  Standardisierung an und ist im erzeugten Artefakt zu prüfen.

System-generierte Felder werden erst dann in die erzeugte validierte
Arbeitsmappe geschrieben, wenn die Prüfung ohne Blocking Errors abgeschlossen
ist (`pipeline_valid = true`). Solange Blocking Errors bestehen, wird keine
validierte Artefakt-Arbeitsmappe mit diesen Rückschreibungen erzeugt.

Die Ordner `Validation_output/` und lokal erzeugte Prüfberichte sind
Laufzeitartefakte und werden grundsätzlich nicht committed. Massgeblich sind
die von GitHub Actions bereitgestellten Berichte und Artefakte des jeweiligen
Laufs.

## 7. Änderungen über GitHub

Änderungen sind nachvollziehbar und auf den erforderlichen Umfang beschränkt
einzureichen. Ein Beitrag soll insbesondere:

- nur die vorgesehenen Data-Dictionary-Dateien und unmittelbar erforderliche
  Begleitdokumentation enthalten;
- keine Zugangsdaten, personenbezogenen Daten oder nicht freigegebenen
  vertraulichen Inhalte enthalten;
- im Pull Request Organisation, Data-Dictionary-Owner, Version, Status und
  Zweck der Änderung nennen;
- vorhandene Dateien nicht unbeabsichtigt überschreiben oder löschen;
- vor dem Zusammenführen die automatisierten Prüfungen durchlaufen.

Der Hauptbranch darf nur über einen kontrollierten Review-Prozess geändert
werden. Die technische Annahme eines Pull Requests bestätigt ausschliesslich,
dass die Repository- und Validierungsregeln eingehalten wurden. Sie ist keine
fachliche Freigabe des Data Dictionary.

### 7.1 Standardweg für externe Organisationen

Externe Organisationen erhalten grundsätzlich keine Mitgliedschaft und keinen
direkten Schreibzugriff auf das Original-Repository. Sie erstellen einen Fork,
legen ihre Arbeitsmappe dort unter
`WIP data dictionaries/<Organisation>/` ab und reichen die Änderung über einen
Pull Request ein.

Der Pull Request muss mindestens folgende Angaben enthalten:

- verantwortliche Organisation;
- Data-Dictionary-Owner;
- Bezeichnung und Version des Data Dictionary;
- gewünschter Status, bei einer Ersteinreichung grundsätzlich `WIP`;
- Zweck und Umfang der Änderung;
- Bestätigung, dass die Organisation die Verantwortung für Inhalt, Rechte,
  Qualitätssicherung und Freigabe trägt.

Das Zusammenführen erfolgt ausschliesslich durch die Repository-Verwaltung nach
Prüfung des Änderungsumfangs und der automatisierten Statuschecks.

### 7.2 Fallback per E-Mail

Kann eine Organisation den Fork-und-Pull-Request-Prozess nicht selbst
durchführen, darf sie eine Anfrage an **info@intop3.ch** senden.

Das Postfach wird einmal pro Woche bearbeitet. Es kann deshalb zu einer
Wartezeit von bis zu sieben Kalendertagen kommen. Die Anfrage muss die oben
genannten Angaben sowie die auszufüllende oder bereits ausgefüllte
Strukturvorlage enthalten.

Die Einreichung per E-Mail begründet keinen Anspruch auf Aufnahme, Bearbeitung
innerhalb einer kürzeren Frist oder fachliche Freigabe. Die inhaltliche
Verantwortung verbleibt auch bei diesem Weg vollständig beim
Data-Dictionary-Owner.

## 8. Technische Repository-Bereiche

- `templates/`: verbindliche leere Ausgangsvorlagen;
- `scripts/`: produktive Validierungs- und Hilfsprogramme;
- `validator_tests/`: automatisierte Regressionstests des Validators;
- `docs/`: Anleitungen und verbindliche Repository-Regeln;
- `Validation_output/`: lokal erzeugte, nicht zu committende Laufzeitartefakte.

Data-Dictionary-Arbeitsmappen dürfen nicht unter `scripts/`, `validator_tests/`
oder `Validation_output/` abgelegt werden.
