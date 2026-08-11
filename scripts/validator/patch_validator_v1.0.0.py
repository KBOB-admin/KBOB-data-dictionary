#!/usr/bin/env python3
"""
Patch validator to support v1.0.0 template columns.
This adds validation for RelatedDocumentName and RelatedDocumentItemReference
columns in Properties, Values, GroupOfProperties, and Data_Template sheets.

The changes are backward compatible with v0.9.5 templates.
"""

import re
from pathlib import Path


def get_patch_content():
    """Returns the patch content to be applied to validate_strukturvorlage.py"""
    
    patch = '''
# Document reference validation helper
def validate_document_reference(self, sheet_name: str, row_idx: int, doc_name: str, 
                                 doc_item_ref: str, document_name_en_set: set,
                                 headers: dict, row, prefix: str = ""):
    """Validate a document reference against the Documents sheet.
    
    Args:
        sheet_name: Name of the sheet being validated
        row_idx: Row index for error reporting
        doc_name: The RelatedDocumentName value
        doc_item_ref: The RelatedDocumentItemReference value
        document_name_en_set: Set of valid document names from Documents sheet
        headers: Header mapping for column lookups
        row: Current row data
        prefix: Optional prefix for column names (e.g., "Related")
    """
    # Validate RelatedDocumentName if present
    if doc_name and document_name_en_set and doc_name not in document_name_en_set:
        self.add('error', 'unknown_related_document_id', 
                 f'{prefix}RelatedDocumentName must reference an existing Documents.DocumentName (EN). Got: {doc_name}', 
                 sheet=sheet_name, row=row_idx)
    
    # Validate RelatedDocumentItemReference format if present
    # Item reference should typically be a section/chapter reference like "3.2.1" or "Clause 4"
    if doc_item_ref and not doc_name:
        self.add('warning', 'document_item_without_document', 
                 f'{prefix}RelatedDocumentItemReference is filled but {prefix}RelatedDocumentName is empty. Document reference recommended.',
                 sheet=sheet_name, row=row_idx)
'''
    return patch


def patch_validate_classes(content: str) -> str:
    """Patch validate_classes method to add RelatedDocumentItemReference support"""
    
    # Find the section where RelatedDocumentName is read and add ItemReference
    old_code = '''            related_document = self._cell(row, headers.get('RelatedDocumentName (EN)', 35 if self.has_public_dictionary() else 34))
            self._require_en_plus_one_local(row, idx, sheet_name, headers.get('Designation (EN)', 15),'''
    
    new_code = '''            related_document = self._cell(row, headers.get('RelatedDocumentName (EN)', 35 if self.has_public_dictionary() else 34))
            related_document_item = self._cell(row, headers.get('RelatedDocumentItemReference', 36 if self.has_public_dictionary() else 35))
            self._require_en_plus_one_local(row, idx, sheet_name, headers.get('Designation (EN)', 15),'''
    
    if old_code in content:
        content = content.replace(old_code, new_code)
        print("  Patched validate_classes: Added RelatedDocumentItemReference")
    
    # Add validation for the document item reference after the existing related_document validation
    old_validation = '''            if related_document and document_name_en_set and related_document not in document_name_en_set:
                self.add('error', 'unknown_related_document_id', f'RelatedDocumentName (EN) must reference an existing Documents.DocumentName (EN). Got: {related_document}', sheet=sheet_name, row=idx)'''
    
    new_validation = '''            if related_document and document_name_en_set and related_document not in document_name_en_set:
                self.add('error', 'unknown_related_document_id', f'RelatedDocumentName (EN) must reference an existing Documents.DocumentName (EN). Got: {related_document}', sheet=sheet_name, row=idx)
            if related_document_item and not related_document:
                self.add('warning', 'document_item_without_document', 'RelatedDocumentItemReference is filled but RelatedDocumentName (EN) is empty. Document reference recommended.', sheet=sheet_name, row=idx)'''
    
    if old_validation in content:
        content = content.replace(old_validation, new_validation)
        print("  Patched validate_classes: Added item reference validation")
    
    return content


def patch_validate_properties(content: str) -> str:
    """Patch validate_properties method to add document reference support"""
    
    # Find where Provenance is read and add document columns
    old_code = '''            prov = self._cell(row, headers.get('Provenance (PROV)', 34))
            self._require_en_plus_one_local(row, idx, sheet_name, headers.get('Designation (EN)', 11),'''
    
    new_code = '''            prov = self._cell(row, headers.get('Provenance (PROV)', 34))
            related_document = self._cell(row, headers.get('RelatedDocumentName (EN)', 36))
            related_document_item = self._cell(row, headers.get('RelatedDocumentItemReference', 37))
            self._require_en_plus_one_local(row, idx, sheet_name, headers.get('Designation (EN)', 11),'''
    
    if old_code in content:
        content = content.replace(old_code, new_code)
        print("  Patched validate_properties: Added document reference columns")
    
    # Add validation for document references - find the end of the property validation loop
    # Insert before the final _check_unique_labels call
    old_end = '''            if not prov:
                self.add('error', 'missing_property_provenance', 'Properties.Provenance (PROV) is required.', sheet=sheet_name, row=idx)
        self._check_unique_labels(seen_labels, sheet_name, 'duplicate_property_label', 'Property label')'''
    
    new_end = '''            if not prov:
                self.add('error', 'missing_property_provenance', 'Properties.Provenance (PROV) is required.', sheet=sheet_name, row=idx)
            # Validate document references for v1.0.0+ templates
            if related_document:
                document_name_en_set = getattr(self, '_document_name_en_set', None)
                if document_name_en_set is None:
                    document_name_en_set = set()
                    documents_sheet = self._documents_sheet()
                    if documents_sheet in self.wb.sheetnames:
                        documents_ws = self.wb[documents_sheet]
                        documents_headers = self._sheet_headers(documents_sheet)
                        for doc_idx, doc_row in self._iter_data_rows(documents_ws, self._sheet_start_row(documents_sheet)):
                            doc_name_en = self._cell(doc_row, documents_headers.get('DocumentName (EN)', 7))
                            if doc_name_en:
                                document_name_en_set.add(doc_name_en)
                    self._document_name_en_set = document_name_en_set
                if related_document not in document_name_en_set:
                    self.add('error', 'unknown_related_document_id', f'RelatedDocumentName (EN) must reference an existing Documents.DocumentName (EN). Got: {related_document}', sheet=sheet_name, row=idx)
            if related_document_item and not related_document:
                self.add('warning', 'document_item_without_document', 'RelatedDocumentItemReference is filled but RelatedDocumentName (EN) is empty. Document reference recommended.', sheet=sheet_name, row=idx)
        self._check_unique_labels(seen_labels, sheet_name, 'duplicate_property_label', 'Property label')'''
    
    if old_end in content:
        content = content.replace(old_end, new_end)
        print("  Patched validate_properties: Added document reference validation")
    
    return content


def patch_validate_values(content: str) -> str:
    """Patch validate_wertekatalog to add document reference support"""
    
    # Find where Provenance is read and add document columns
    old_code = '''            if not prov:
                self.add('error', 'missing_value_provenance', 'Values.Provenance (PROV) is required.', sheet=sheet_name, row=idx)'''
    
    new_code = '''            if not prov:
                self.add('error', 'missing_value_provenance', 'Values.Provenance (PROV) is required.', sheet=sheet_name, row=idx)
            # Validate document references for v1.0.0+ templates
            related_document = self._cell(row, headers.get('RelatedDocumentName (EN)', 21))
            related_document_item = self._cell(row, headers.get('RelatedDocumentItemReference', 22))
            if related_document:
                document_name_en_set = getattr(self, '_document_name_en_set', None)
                if document_name_en_set is None:
                    document_name_en_set = set()
                    documents_sheet = self._documents_sheet()
                    if documents_sheet in self.wb.sheetnames:
                        documents_ws = self.wb[documents_sheet]
                        documents_headers = self._sheet_headers(documents_sheet)
                        for doc_idx, doc_row in self._iter_data_rows(documents_ws, self._sheet_start_row(documents_sheet)):
                            doc_name_en = self._cell(doc_row, documents_headers.get('DocumentName (EN)', 7))
                            if doc_name_en:
                                document_name_en_set.add(doc_name_en)
                    self._document_name_en_set = document_name_en_set
                if related_document not in document_name_en_set:
                    self.add('error', 'unknown_related_document_id', f'RelatedDocumentName (EN) must reference an existing Documents.DocumentName (EN). Got: {related_document}', sheet=sheet_name, row=idx)
            if related_document_item and not related_document:
                self.add('warning', 'document_item_without_document', 'RelatedDocumentItemReference is filled but RelatedDocumentName (EN) is empty. Document reference recommended.', sheet=sheet_name, row=idx)'''
    
    if old_code in content:
        content = content.replace(old_code, new_code)
        print("  Patched validate_wertekatalog: Added document reference validation")
    
    return content


def patch_validate_groups(content: str) -> str:
    """Patch validate_merkmalsgruppenkatalog to add document reference support"""
    
    # Find the loop in validate_merkmalsgruppenkatalog and add document validation
    # Look for the section where provenance is checked
    old_code = '''            if not prov:
                self.add('error', 'missing_group_provenance', 'GroupOfProperties.Provenance (PROV) is required.', sheet=sheet_name, row=idx)'''
    
    new_code = '''            if not prov:
                self.add('error', 'missing_group_provenance', 'GroupOfProperties.Provenance (PROV) is required.', sheet=sheet_name, row=idx)
            # Validate document references for v1.0.0+ templates
            related_document = self._cell(row, headers.get('RelatedDocumentName (EN)', 16))
            related_document_item = self._cell(row, headers.get('RelatedDocumentItemReference', 17))
            if related_document:
                document_name_en_set = getattr(self, '_document_name_en_set', None)
                if document_name_en_set is None:
                    document_name_en_set = set()
                    documents_sheet = self._documents_sheet()
                    if documents_sheet in self.wb.sheetnames:
                        documents_ws = self.wb[documents_sheet]
                        documents_headers = self._sheet_headers(documents_sheet)
                        for doc_idx, doc_row in self._iter_data_rows(documents_ws, self._sheet_start_row(documents_sheet)):
                            doc_name_en = self._cell(doc_row, documents_headers.get('DocumentName (EN)', 7))
                            if doc_name_en:
                                document_name_en_set.add(doc_name_en)
                    self._document_name_en_set = document_name_en_set
                if related_document not in document_name_en_set:
                    self.add('error', 'unknown_related_document_id', f'RelatedDocumentName (EN) must reference an existing Documents.DocumentName (EN). Got: {related_document}', sheet=sheet_name, row=idx)
            if related_document_item and not related_document:
                self.add('warning', 'document_item_without_document', 'RelatedDocumentItemReference is filled but RelatedDocumentName (EN) is empty. Document reference recommended.', sheet=sheet_name, row=idx)'''
    
    if old_code in content:
        content = content.replace(old_code, new_code)
        print("  Patched validate_merkmalsgruppenkatalog: Added document reference validation")
    
    return content


def apply_patches(validator_path: Path):
    """Apply all patches to the validator"""
    
    print(f"Reading validator from: {validator_path}")
    content = validator_path.read_text(encoding='utf-8')
    original_content = content
    
    # Apply patches
    content = patch_validate_classes(content)
    content = patch_validate_properties(content)
    content = patch_validate_values(content)
    content = patch_validate_groups(content)
    
    # Write back if changed
    if content != original_content:
        backup_path = validator_path.with_suffix('.py.v095.backup')
        print(f"Creating backup at: {backup_path}")
        backup_path.write_text(original_content, encoding='utf-8')
        
        print(f"Writing patched validator to: {validator_path}")
        validator_path.write_text(content, encoding='utf-8')
        print("\n✓ Validator patched successfully!")
        print("  Backup saved as: validate_strukturvorlage.py.v095.backup")
    else:
        print("\nNo changes made (validator may already be patched)")


def main():
    import sys
    validator_path = Path(__file__).parent / "validate_strukturvorlage.py"
    if not validator_path.exists():
        # Try to find it relative to repo root
        validator_path = Path(__file__).parent.parent.parent / "scripts" / "validator" / "validate_strukturvorlage.py"
    
    if not validator_path.exists():
        print(f"Error: Could not find validator at {validator_path}")
        sys.exit(1)
    
    apply_patches(validator_path)


if __name__ == "__main__":
    main()
