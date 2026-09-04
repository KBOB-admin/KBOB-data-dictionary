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
- jede befuellte `Data_Template`-Zeile muss in Spalte B eine stabile, kleingeschriebene und mit Bindestrichen strukturierte `DataTemplate-ID` enthalten,
- dieselbe `DataTemplate-ID` darf ueber mehrere Zeilen wiederholt werden und bildet dadurch ein Data Template mit `0..n` unterschiedlichen `GroupOfProperties`,
- alle Zeilen derselben `DataTemplate-ID` muessen auf dieselbe Class verweisen und identische Governance-Werte verwenden; unterschiedliche LOIN-Werte duerfen als mehrere kontextuelle Angaben aggregiert werden,
- dieselbe Property darf innerhalb derselben Kombination aus `DataTemplate-ID` und GoP-Kontext nicht doppelt zugeordnet werden,
- Referenzen auf Wertelisten müssen auf vorhandene `Values` zeigen,
- Dokumentreferenzen müssen – wo vorgesehen – auf vorhandene `Documents` verweisen.

Wenn eine Referenz nicht aufgelöst werden kann, entsteht ein Fehler.

## F. Formale Prüfungen

Bestimmte Felder werden zusätzlich formal geprüft.

Dazu gehören zum Beispiel:

- URI-Formatprüfungen,
- `Classes.IFC_URI` und IFC-Entity-Felder werden für echte IFC-ausgerichtete Klassen geprüft; lokale Dokumenttyp-Taxonomieklassen, die als `Document type taxonomy` markiert sind, dürfen diese Felder leer lassen, weil sie keine exakten IFC-Objekt-/TypeObject-Entsprechungen darstellen,
- `Header.DictionaryUri` verwendet vor einer Publikationsfreigabe eine neutrale, aus `OrganizationCode` und `DictionaryCode` abgeleitete `https://example.com/...`-URI,
- eine LINDAS-URI wird niemals allein aus Workbook-Metadaten zugeteilt; sie wird erst nach einer expliziten Publikationsentscheidung eingetragen,
- eine freigegebene LINDAS-URI muss die konkrete SemVer-Version am Ende enthalten und mit `Header.DictionaryVersion` übereinstimmen,
- eine explizit freigegebene LINDAS-URI wird als autoritative Override-URI unverändert akzeptiert; der Validator leitet dafür keinen alternativen Namespace ab,
- andere URI-Authorities sind für `Header.DictionaryUri` nicht zulässig,
- `Header.I14yDatasetUri` ist optional und muss, falls gesetzt, eine absolute `i14y.admin.ch`-Dataset-URI sein; sie ist Metadatum und keine Dictionary-Identität,
- Identifier- oder Code-Formate,
- kontrollierte Werte aus Listen,
- Konsistenz zwischen Feldern und Referenzen.

## G. Reproduzierbare Referenzartefakte

Der Validator verwendet für externe Referenzprüfungen ausschliesslich Artefakte, die im Repository liegen:

- `resources/qudt/units.ttl` für `Properties.QUDT URI`
- `resources/bsdd/ifc4.3-uri-cache.json` für IFC-/bSDD-Identifier
- `resources/ifc/ifc4x3-add2-object-type-pairs.json` für schema-konforme Paare aus IFC-Objekt- und TypeObject-Entitäten

Damit darf die GitHub-Validierung nicht von lokalen Dateien unter `/home/...` oder anderen persönlichen Arbeitsverzeichnissen abhängen.

Lernpunkt aus `unknown_qudt_unit_uri` für `http://qudt.org/vocab/unit/M2`: Die URI ist gültig und im QUDT-Referenzartefakt vorhanden. Der Fehler entstand, weil die frühere Validator-Konfiguration auf eine lokale QUDT-Datei ausserhalb des Repositories zeigte. In einer reproduzierbaren GitHub-Umgebung war diese Datei nicht garantiert vorhanden; dadurch wurde die lokale Referenzmenge leer oder abweichend aufgebaut und `M2` konnte fälschlich als unbekannt erscheinen.

### IFC-Mapping in `Classes`

Die vier IFC-Felder beschreiben unterschiedliche Aspekte des Mappings:

- `IFC_URI` verweist auf den bSDD-Identifier der gemappten IFC-Klasse, gegebenenfalls einschliesslich PredefinedType, zum Beispiel `.../class/IfcTankVESSEL`.
- `IfcObject Entity` enthält die IFC-Entität der Objektebene, zum Beispiel `IfcTank`.
- `IfcTypeObject Entity` ist optional und dokumentiert, ob das Mapping zusätzlich für die zugehörige IFC-Typebene gilt.
- `PredefinedType` enthält, falls verwendet, den kontrollierten IFC-Enumerationswert, zum Beispiel `VESSEL`.

Für `IfcTypeObject Entity` gilt ausdrücklich:

- **Leer = nur Objektebene.** Ein leerer Wert ist gültig und erzeugt keine Warnung.
- **Ausgefüllt = Mapping gilt auch für die angegebene Typebene.** Dann muss die Typentität zum `IfcObject Entity` passen. Für `IfcTank` ist ausschliesslich `IfcTankType` zulässig.

Der PredefinedType ist eine separate Aussage und wird unabhängig von der optionalen TypeObject-Spalte geprüft. Beispiel eines spezialisierten Mappings:

```text
IFC_URI:              https://identifier.buildingsmart.org/uri/buildingsmart/ifc/4.3/class/IfcTankVESSEL
IfcObject Entity:     IfcTank
IfcTypeObject Entity: IfcTankType
PredefinedType:       VESSEL
```

`IfcTankVESSEL` ist hierbei ein bSDD-Identifier für die Kombination aus Objektentität und PredefinedType. Es ist keine IFC-TypeObject-Entität und darf deshalb nicht in `IfcTypeObject Entity` eingetragen werden. Ebenso ist `IfcTankVessel` keine gültige IFC-Entität.

Die Prüfungen bleiben bewusst getrennt:

1. `IfcObject Entity` und `IfcTypeObject Entity` müssen ein schema-konformes Paar bilden, sofern die optionale TypeObject-Spalte befüllt ist.
2. `PredefinedType` muss ein für die Objektentität zulässiger IFC-Wert sein.
3. `IFC_URI` muss mit der Kombination aus `IfcObject Entity` und `PredefinedType` übereinstimmen, wenn ein PredefinedType angegeben ist.

### Mehrere IFC-Pset-/Qto-Referenzen in `Properties`

Die Spalte `IfcPropertySet (Pset) / IfcQuantitySet (Qto)` akzeptiert genau zwei Eingabeformen:

- **Einzelwert:** ein einzelner Set-Name oder eine einzelne zulässige absolute IRI, zum Beispiel `Pset_PipeSegmentTypeCommon`.
- **Mehrfachwert:** eine syntaktisch gültige JSON-Liste aus nicht leeren Strings, zum Beispiel `["Pset_PipeSegmentTypeCommon", "Qto_PipeSegmentBaseQuantities"]`.

Zeilenumbruch-, Semikolon- oder einfache Kommalisten sind nicht zulässig. JSON-Listen verwenden doppelte Anführungszeichen. Jeder Listeneintrag wird einzeln nach denselben Regeln wie ein Einzelwert validiert. Leere Einträge, Nicht-String-Werte und Duplikate sind Fehler.

Diese Konvention entspricht der JSON-Listensyntax in `Values.Enumeration (EN)` und bei Mehrfachwerten in `Data_Template`. Benutzer müssen damit nur zwischen einem Einzelwert und einer JSON-Liste unterscheiden.

## H. Ergebnisse

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
