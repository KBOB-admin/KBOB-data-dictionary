# Schritt-für-Schritt-Anleitung zur Strukturvorlage

## 1. Zweck dieser Anleitung

Diese Anleitung beschreibt in einfacher Form, wie eine Organisation ein Data
Dictionary mit der Strukturvorlage erstellt, prüft und als Beitrag zum
Repository einreicht.

Die technische Validierung prüft die Struktur und maschinenlesbare Regeln. Sie
prüft nicht, ob die fachlichen Inhalte richtig, vollständig oder für einen
bestimmten Zweck geeignet sind. Dafür bleibt ausschliesslich der jeweilige
Data-Dictionary-Owner verantwortlich.

## 2. Vor dem Start

Benötigt werden:

- eine benannte verantwortliche Organisation;
- ein benannter Data-Dictionary-Owner;
- ein GitHub-Konto für den Standardweg;
- Microsoft Excel oder ein kompatibles Tabellenprogramm;
- Klarheit über Zweck, Inhalt und geplante Version des Data Dictionary.

Lesen Sie vor der Bearbeitung die
[Housekeeping-Regeln](repository-housekeeping.md).

## 3. Aktuelle Strukturvorlage beziehen

1. Öffnen Sie den Ordner `templates/`.
2. Verwenden Sie ausschliesslich die aktuell bezeichnete leere Vorlage:
   `2026_09_Strukturvorlage_Data_Dictionary_leer_v1.1.0.xlsx`.
3. Laden Sie die Datei herunter oder kopieren Sie sie in Ihren eigenen
   Arbeitsbereich.
4. Verändern oder überschreiben Sie die Originaldatei unter `templates/`
   nicht.

Verwenden Sie für neue Arbeiten keine ältere Vorlage aus einem WIP-,
Published- oder Archivordner.

## 4. Datei eindeutig benennen

Der Dateiname soll Organisation, Data Dictionary und Version erkennen lassen.
Empfohlenes Schema:

`<Organisation>_<Data-Dictionary-Name>_v<MAJOR.MINOR.PATCH>.xlsx`

Beispiel:

`MusterAG_Facility-Management_v0.1.0.xlsx`

Eine bereits publizierte Versionsnummer darf nicht mit verändertem Inhalt
erneut verwendet werden.

## 5. Strukturvorlage ausfüllen

Beachten Sie insbesondere:

1. Benennen Sie vorhandene Tabellenblätter nicht um.
2. Verschieben oder löschen Sie keine Kopfzeilen und Strukturblöcke.
3. Füllen Sie die erforderlichen Angaben im Blatt `Header` aus.
4. Erfassen Sie Klassen im Blatt `Classes`.
5. Erfassen Sie Properties im Blatt `Properties`.
6. Erfassen Sie globale Wertelisten im Blatt `Values`.
7. Erfassen Sie Quellen und Referenzdokumente im Blatt `Documents`.
8. Verwenden Sie `GroupOfProperties` und `Data_Template` entsprechend der
   vorgesehenen Struktur.
9. Verwenden Sie für Dropdown-Felder nur Werte aus dem Blatt `Rules`.
10. Tragen Sie nachvollziehbare Versions-, Status- und Provenance-Angaben ein.

Fachliche Definitionen, Benennungen, Übersetzungen, Zuordnungen und
Wertelisten müssen durch den Data-Dictionary-Owner geprüft werden.

## 6. Bedeutung system-generierter Felder

Einige Felder werden durch den Validierungsprozess sicher abgeleitet oder
normalisiert. Diese Felder werden nicht sofort in die eingereichte Datei
geschrieben.

Eine validierte Arbeitsmappe mit system-generierten Rückschreibungen wird erst
erzeugt, wenn keine Blocking Errors mehr vorhanden sind und
`pipeline_valid = true` ist.

Solange Blocking Errors bestehen:

- wird ein Prüfbericht erzeugt;
- werden die Fehler und Warnungen angezeigt;
- wird keine validierte Artefakt-Arbeitsmappe mit system-generierten Feldern
  bereitgestellt.

## 7. Eigenen WIP-Ordner vorbereiten

Neue oder noch nicht freigegebene Dateien gehören nach:

`WIP data dictionaries/<Organisation>/`

Verwenden Sie einen kurzen und eindeutig erkennbaren Organisationsnamen. Legen
Sie keine Arbeitsmappen unter `templates/`, `scripts/`, `validator_tests/` oder
`Validation_output/` ab.

Wenn der Organisationsordner noch nicht existiert, wird er im eigenen Fork
angelegt. Bei Nutzung der GitHub-Weboberfläche kann zuerst eine kleine
`README.md` im neuen Organisationsordner erstellt und danach die `.xlsx`-Datei
dorthin hochgeladen werden.

## 8. Standardweg: Beitrag über Fork und Pull Request

### 8.1 Repository forken

1. Melden Sie sich bei GitHub an.
2. Öffnen Sie das Repository `KBOB-admin/KBOB-data-dictionary`.
3. Wählen Sie oben rechts `Fork`.
4. Erstellen Sie den Fork unter Ihrem eigenen GitHub-Konto oder Ihrer
   Organisation.

Der Fork ist Ihre eigene Arbeitskopie. Dadurch erhalten Sie keinen direkten
Schreibzugriff auf das Original-Repository.

### 8.2 Arbeitsdatei im Fork ablegen

1. Öffnen Sie in Ihrem Fork den Ordner `WIP data dictionaries`.
2. Öffnen oder erstellen Sie den Unterordner Ihrer Organisation.
3. Laden Sie ausschliesslich die vorgesehene `.xlsx`-Datei und gegebenenfalls
   eine kurze organisationsbezogene `README.md` hoch.
4. Prüfen Sie vor dem Commit den vollständigen Zielpfad und Dateinamen.
5. Beschreiben Sie den Commit kurz und sachlich, beispielsweise:
   `Add MusterAG Facility Management v0.1.0 as WIP`.

### 8.3 Pull Request erstellen

1. Wählen Sie in Ihrem Fork `Contribute` und danach `Open pull request`.
2. Ziel ist der Branch `main` des Original-Repositorys.
3. Verwenden Sie einen eindeutigen Titel.
4. Füllen Sie die bereitgestellte Pull-Request-Vorlage vollständig aus.
5. Bestätigen Sie die Verantwortung des Data-Dictionary-Owners.
6. Erstellen Sie den Pull Request.

### 8.4 Automatische Prüfung abwarten

Nach dem Erstellen des Pull Requests laufen die automatisierten Prüfungen.

- Ein grüner technischer Check bedeutet, dass die automatisierten Regeln
  bestanden wurden.
- Ein roter Check bedeutet, dass mindestens eine technische Prüfung nicht
  bestanden wurde.
- Ein bestandener Check ist keine fachliche Freigabe.

Beheben Sie gemeldete Blocking Errors in Ihrem Fork und aktualisieren Sie den
Pull Request. Warnungen sind durch den Data-Dictionary-Owner zu prüfen.

### 8.5 Review und Zusammenführung

Die Repository-Verwaltung prüft:

- den richtigen Zielordner;
- Dateiname und Version;
- den Umfang der Änderung;
- die Ergebnisse der automatisierten Prüfungen;
- die Vollständigkeit der Pull-Request-Angaben.

Die Repository-Verwaltung prüft nicht die fachliche Richtigkeit des gesamten
Data Dictionary. Erst nach erfolgreichem Review kann sie den Pull Request
zusammenführen.

## 9. Fallback: Anfrage per E-Mail

Kann Ihre Organisation den Fork-und-Pull-Request-Prozess nicht selbst
durchführen, senden Sie eine Anfrage an:

**info@intop3.ch**

Das Postfach wird einmal pro Woche bearbeitet. Es kann deshalb zu einer
Wartezeit von bis zu sieben Kalendertagen kommen.

Die E-Mail soll enthalten:

- Name der Organisation;
- Name oder Funktion des Data-Dictionary-Owners;
- Bezeichnung und Version des Data Dictionary;
- kurze Beschreibung des Zwecks;
- gewünschte Ablage, bei einer Ersteinreichung grundsätzlich `WIP`;
- die ausgefüllte Strukturvorlage als `.xlsx`;
- Bestätigung, dass die Organisation die Verantwortung für Inhalt, Rechte,
  Qualitätssicherung und Freigabe trägt.

Die E-Mail begründet keinen Anspruch auf Aufnahme oder fachliche Prüfung. Die
inhaltliche Verantwortung verbleibt vollständig bei der einreichenden
Organisation und ihrem Data-Dictionary-Owner.

## 10. Wechsel von WIP zu SHARED

Ein Wechsel nach `SHARED` ist erst vorgesehen, wenn:

- keine Blocking Errors mehr bestehen;
- alle Warnungen geprüft wurden;
- Owner, Version und Provenance vollständig sind;
- der Data-Dictionary-Owner die konkrete Version zur Koordination freigegeben
  hat.

Der Statuswechsel erfolgt über einen eigenen Pull Request und muss begründet
werden.

## 11. Wechsel nach PUBLISHED

Ein Wechsel nach `PUBLISHED` erfordert zusätzlich:

- abgeschlossene fachliche Qualitätssicherung;
- ausdrückliche Freigabe der konkreten Version durch den
  Data-Dictionary-Owner;
- dokumentierte Version, Verantwortlichkeit, Status und Provenance;
- Übereinstimmung zwischen freigegebener und validierter Datei.

Ein technisch bestandener Validierungslauf allein reicht nicht aus.

## 12. Archivierung

Abgelöste oder zurückgezogene Versionen werden nach
`ARCHIVED data dictionaries/<Organisation>/` verschoben. Sie bleiben für die
Nachvollziehbarkeit erhalten, dürfen aber nicht als aktuelle Version verwendet
werden.

## 13. Kurzcheck vor einer Einreichung

Prüfen Sie vor dem Pull Request oder der E-Mail:

- aktuelle leere Vorlage verwendet;
- richtige Organisation und richtige Version im Dateinamen;
- Datei unter `WIP data dictionaries/<Organisation>/` abgelegt;
- keine temporären Dateien wie `~$...xlsx` enthalten;
- Header, Owner, Version, Status und Provenance ausgefüllt;
- keine vertraulichen oder unzulässigen personenbezogenen Daten enthalten;
- Verantwortungserklärung vorbereitet;
- Validierungsbericht gelesen und Blocking Errors verstanden.

## 14. Verantwortungsabgrenzung

Der Data-Dictionary-Owner bleibt verantwortlich für Inhalt, fachliche Qualität,
Rechte, Freigabe und Verwendung des Data Dictionary.

Die Repository-Verwaltung und der Validierungsdienst stellen eine technische
Ablage- und Prüfstruktur bereit. Sie übernehmen keine fachliche Gewährleistung
für die eingereichten Inhalte.
