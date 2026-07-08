# Validierungslogik

Diese Datei erklärt in einfacher Sprache, was der Validator im öffentlichen MVP prüft.

## A. Struktur der Vorlage

Der Validator prüft zuerst die Struktur der Arbeitsmappe.

Dazu gehört insbesondere:

- die erforderlichen Tabellenblätter müssen vorhanden sein,
- die öffentlichen Blattnamen dürfen nicht verändert werden,
- wichtige Spaltennamen dürfen nicht verändert werden,
- Hinweiszeilen und Kopfzeilen dürfen nicht gelöscht oder verschoben werden,
- die Datei muss technisch lesbar bleiben.

Die aktuellen Blattnamen im MVP-Kern sind:

- `Header`
- `Classes`
- `Properties`
- `Values`
- `Documents`
- `GroupOfProperties`
- `Rules`
- `Data_Template`

Zusätzlich enthalten alle aktuellen Vorlagen auch:

- `Impressum`
- `Anleitung`
- `Gelöschte Bauteile`

Die Public-Varianten enthalten ausserdem zusätzlich:

- `Dictionary_public`

Wenn die Struktur nicht stimmt, kann der Validator die Datei nicht zuverlässig interpretieren.

Wichtig für die Weiterentwicklung: Änderungen an Blattnamen, Kernspalten oder Guidance müssen immer konsistent über alle vier kanonischen `.xlsx`-Dateien umgesetzt und gleichzeitig in der Validator-Logik nachvollzogen werden. Der Validator darf dabei keine veralteten Zielstrukturen künstlich als primären Soll-Zustand konservieren.

## B. Pflichtfelder

Bestimmte Felder sind Pflichtfelder.

Wenn Pflichtangaben fehlen, meldet der Validator Fehler. Beispiele:

- Pflichtwerte in `Header`
- Pflichtangaben in `Classes`, `Properties`, `Documents` oder `GroupOfProperties`
- fehlende Referenzen in `Data_Template`

Die Fehlermeldung zeigt an, welche Zelle oder welche Zeile ergänzt werden muss.

## C. Datentypen und Datumsformat

Der Validator prüft, ob Werte zum erwarteten Typ passen.

Beispiele:

- Text
- Zahl
- Boolean
- Datum/Zeit
- URI
- kontrollierte Werte aus Dropdowns

Datumsfelder müssen im MVP als ISO-8601-Datum mit Zeitzone vorliegen, zum Beispiel:

- `2026-06-18T15:30+02:00`

## D. Dropdowns / erlaubte Werte

Viele Spalten haben erlaubte Werte oder Dropdown-Listen.

Dann gilt:

- nur diese Werte sind erlaubt,
- freie Texteingaben können fehlschlagen,
- die Werte müssen zur jeweiligen Liste in `Rules` passen.

Wichtige aktuelle Beispiele sind:

- `Classes.Class-Assignment` muss zur Liste in `Rules.Class-Assignment` passen,
- `Properties.Property-Assignment` muss zur Liste in `Rules.Property-Assignment` passen,
- die englischen Werte in `Rules.Class-Assignment` sind die massgebliche Dropdown-Liste für die Klassenzuordnung,
- die englischen Werte in `Rules.Property-Assignment` sind die massgebliche Dropdown-Liste für die Merkmals-Zuordnung,
- `Classes.Klassen-Zuordnung (DE)`, `Classes.Affectation de classe (FR)` und `Classes.Assegnazione della classe (IT)` werden system-generiert aus `Classes.Class-Assignment` plus den zugehörigen `Rules`-Übersetzungsspalten,
- `Properties.Merkmals-Zuordnung (DE)`, `Properties.Attribution de propriété (FR)` und `Properties.Assegnazione della proprietà (IT)` werden system-generiert aus `Properties.Property-Assignment` plus den zugehörigen `Rules`-Übersetzungsspalten.

## E. Referenzen zwischen Tabellen

Der aktuelle MVP verwendet ausserdem die neue Benennung der Zuordnungsfelder:

- `Class-Assignment` statt `Classification`
- `Property-Assignment` statt `Property Classification`

Diese Benennung ist Teil der aktuellen synchronisierten Vorlagenfamilie und muss in Vorlage, Beispiel und Validator konsistent bleiben.

## E. Referenzen zwischen Tabellen

Der Validator prüft Verknüpfungen zwischen den Tabellen.

Beispiele:

- `Data_Template` muss auf vorhandene `Classes` und `Properties` verweisen,
- Referenzen auf Wertelisten müssen auf vorhandene `Values` zeigen,
- Dokumentreferenzen müssen – wo vorgesehen – auf vorhandene `Documents` verweisen.

Wenn eine Referenz nicht aufgelöst werden kann, entsteht ein Fehler.

## F. Formale Prüfungen

Bestimmte Felder werden zusätzlich formal geprüft.

Dazu gehören zum Beispiel:

- URI-Formatprüfungen,
- Identifier- oder Code-Formate,
- kontrollierte Werte aus Listen,
- Konsistenz zwischen Feldern und Referenzen.

## G. Ergebnisse

Der Validator liefert drei Arten von Resultaten:

### Fehler
Müssen behoben werden. Die Datei ist in diesem Zustand nicht gültig.

### Warnungen
Sollten überprüft werden. Die Datei ist möglicherweise unvollständig oder missverständlich.

### Normalisierungen
Zeigen an, dass der Validator einen Wert automatisch interpretiert, abgeleitet oder standardisiert hat.

## Wichtiger Hinweis zum aktuellen MVP

Der öffentliche MVP konzentriert sich auf die **Validierung von ausgefüllten Arbeitsmappen**.

- Die Validierung ausgefüllter `.xlsx`-Dateien ist der eigentliche Nutzer-Workflow.
- Die Validierung der leeren Vorlage ist nur der Standard-Smoke-Test der GitHub Action.

Nicht Teil des aktuellen öffentlichen MVP sind:

- RDF-Erzeugung
- bSDD-Publikation
- i14y-Publikation
- LINDAS-Publikation
- sonstige Export-/Publishing-Prozesse
