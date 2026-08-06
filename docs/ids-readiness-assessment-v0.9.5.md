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

## Revised Verdict

Option C applies for reusable Data Dictionary templates, with a separate IDS Creator needed for project-specific IDS composition.

The current templates can support a useful IDS starting-point subset, but they should not be extended with global project-specific requirement cardinality fields unless a publisher intentionally wants to publish normative IDS defaults.

`Data Dictionary` and `Data Template` resources should remain reusable governed source artefacts. `Required`, `Optional` and `Prohibited` are normally user-, project- or exchange-specific choices and should be authored in a separate IDS working configuration.

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

## Reclassified Gaps

The current templates do not explicitly publish project-specific IDS requirement choices:

- required versus optional semantics;
- prohibited semantics;
- a clear authoring rule that maps matrix `x` to required or optional.

This is no longer treated as a blocking Data Template defect. A matrix `x` means "included in this reusable Data Template", not "required in every IDS".

The current reader sets `DDClassProperty.is_required=True` for every populated matrix cell. That is an implementation default and must not be treated as a normative IDS cardinality source.

## Supported IDS Subset Without Template Extension

A generated IDS can safely include:

- IFC entity applicability from Class metadata;
- IFC predefined type applicability where explicitly available;
- Property requirements for included Properties;
- allowed Enumeration value subsets;
- datatype and unit constraints where mapped;
- specification metadata from Data Template, Data Dictionary and source workbook metadata.

This subset is suitable for prepopulating an IDS Creator. It loses reliable distinction between required, optional and prohibited Properties only if an application tries to generate a finished IDS directly from the Data Template without user configuration.

## Paused Template Extension

Do not add these previously proposed global fields to the canonical v0.9.5 Data Dictionary templates:

- `IDS Requirement`
- `Minimum occurrences`
- `Maximum occurrences`

The official buildingSMART IDS XSD `version="1.0.0"` exposes a `cardinality` attribute on requirement facets, with values `required`, `optional` and `prohibited`. It does not expose a general user-facing `minOccurs` / `maxOccurs` property-requirement model for this MVP. Minimum/maximum occurrence controls should therefore not be added unless a future target schema or validator workflow justifies them explicitly.

Normative publisher defaults may still be useful in the future, but they should be optional defaults that the IDS Creator can import and override, not mandatory Data Template semantics.

## Recommendation

Do not change the canonical templates for project-specific IDS requirement cardinality. Build IDS generation through a dedicated IDS Creator:

1. import a `Data Template` as a reusable baseline;
2. copy its Class, IFC applicability, included Properties, Enumeration subsets, datatypes, units and Documents into IDS working state;
3. require the user to choose `required`, `optional` or `prohibited` for each selected requirement;
4. generate and validate `.ids` from that working state.
