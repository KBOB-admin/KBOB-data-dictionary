# IDS Readiness Assessment for v0.9.5 Templates

Date: 2026-07-16
Author: datadict

## Scope

Inspected templates:

- `templates/Strukturvorlage_DataDictionary_empty_public_v0.9.5.xlsx`
- `templates/Strukturvorlage_DataDictionary_empty_v0.9.5.xlsx`

Inspected implementation:

- `P_workspace_codat3_schemaforge_public/scripts/map_to_rdf.py`
- `P_workspace_codat3_schemaforge_public/scripts/export_i14y_json.py`
- `scripts/validator/validate_strukturvorlage.py`
- `scripts/validator/readers/current_template_reader.py`

## Verdict

Option B/C applies.

The current templates can support a useful IDS subset, but they are not sufficient for unambiguous, governance-grade IDS generation without an explicit authoring extension.

## What Is Available

The current authoring model and RDF mapper provide:

- Data Template rows as governed `dd:DataTemplate` resources.
- Applicability through the Class selected in `Data_Template` and the Class-level `IfcObject Entity` / `PredefinedType`.
- Property inclusion through `Data_Template` matrix cells.
- `dd:PropertyRequirement` resources generated per Data Template row and Property.
- Property datatype from `Properties.DataType (Base Type)` and `Properties.DataType (IFC)`.
- Units from `Properties.QUDT URI` and contextual unit predicates where present.
- Enumeration schemes and selected allowed-value subsets.
- Related Documents and LOIN metadata.

## Blocking Gaps

The current templates do not explicitly publish:

- required versus optional semantics;
- minimum cardinality;
- maximum cardinality;
- repeatability;
- prohibited semantics;
- a clear authoring rule that maps matrix `x` to required or optional.

The validator confirms this gap: `validate_matrix()` checks row/property resolvability, allowed-value overrides, Status, Version date and Provenance, but does not validate IDS obligation or cardinality fields because they do not exist.

The current reader sets `DDClassProperty.is_required=True` for every populated matrix cell. That is an implementation default, not a documented authoring rule, and should not be treated as a normative IDS source.

## Supported IDS Subset Without Template Extension

A generated IDS can safely include:

- IFC entity applicability from Class metadata;
- IFC predefined type applicability where explicitly available;
- Property requirements for included Properties;
- allowed Enumeration value subsets;
- datatype and unit constraints where mapped;
- specification metadata from Data Template, Data Dictionary and source workbook metadata.

This subset loses reliable distinction between required and optional Properties, cannot represent repeatability or cardinality, and may overstate requirements if every matrix inclusion is interpreted as mandatory.

## Minimal Template Extension Proposal

Add a small IDS authoring block to `Data_Template`, after the existing property assignment matrix and before `Document - Designation/Bezeichnung/Désignation/Designazione` when document columns are present, otherwise before `LOIN` or `Governance`.

Fields:

| Field | Allowed values | Required | RDF predicate | IDS mapping |
| --- | --- | --- | --- | --- |
| `IDS Requirement` | `Required`, `Optional`, `Prohibited` | Optional, default unset | `dd:idsRequirementLevel` | property facet obligation / generation rule |
| `Minimum occurrences` | integer `0..n` | Optional unless `IDS Requirement = Required` | `dd:minCardinality` | IDS cardinality lower bound where supported |
| `Maximum occurrences` | integer `1..n` or `unbounded` | Optional | `dd:maxCardinality` | IDS cardinality upper bound where supported |

Validation rules:

- `IDS Requirement` must come from `Rules.IDS Requirement`.
- `Required` requires `Minimum occurrences >= 1`.
- `Optional` permits `Minimum occurrences = 0` or blank.
- `Prohibited` requires `Maximum occurrences = 0` or blank with mapper emitting a prohibited facet only if supported.
- `Maximum occurrences` must be greater than or equal to `Minimum occurrences`, except `unbounded`.
- Existing workbooks with blank fields remain valid but IDS readiness is `Ready with limitations` or `Not ready` depending on the generation mode.

Migration impact:

- Existing populated workbooks do not break.
- Existing matrix `x` values keep their current inclusion meaning.
- IDS generation remains disabled or limitation-marked until the new fields are populated for a Data Template.

## Recommendation

Do not change the canonical templates until this proposal is accepted or revised. After acceptance, update both public and non-public templates, validator, reader, RDF mapper, i14y diagnostics, Explorer readiness messages and tests in the same change set.
