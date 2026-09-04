import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts' / 'validator'))

from validate_strukturvorlage import Validator  # noqa: E402


class IfcMappingTests(unittest.TestCase):
    SOURCE = ROOT / 'templates' / 'Strukturvorlage_DataDictionary_empty_v1.0.0.xlsx'

    def setUp(self):
        self.validator = Validator(self.SOURCE)

    def finding_codes(self):
        return {finding.code for finding in self.validator.findings}

    def test_empty_type_object_entity_is_valid_object_only_mapping(self):
        self.validator.validate_ifc_object_type_pair('IfcTank', None, 'Classes', 7)
        self.assertEqual(set(), self.finding_codes())

    def test_ifc_tank_and_ifc_tank_type_are_schema_conformant_pair(self):
        self.validator.validate_ifc_object_type_pair('IfcTank', 'IfcTankType', 'Classes', 7)
        self.assertEqual(set(), self.finding_codes())

    def test_ifc_tank_vessel_is_not_a_type_object_entity(self):
        self.validator.validate_ifc_object_type_pair('IfcTank', 'IfcTankVessel', 'Classes', 7)
        self.assertIn('invalid_ifc_object_type_pair', self.finding_codes())
        self.assertIn('IfcTankType', self.validator.findings[0].message)

    def test_predefined_type_is_validated_separately_from_type_entity(self):
        uri = 'https://identifier.buildingsmart.org/uri/buildingsmart/ifc/4.3/class/IfcTankVESSEL'
        self.validator.validate_predefined_type('IfcTank', 'VESSEL', uri, 'Classes', 7)
        self.assertEqual(set(), self.finding_codes())

    def test_invalid_predefined_type_is_rejected(self):
        uri = 'https://identifier.buildingsmart.org/uri/buildingsmart/ifc/4.3/class/IfcTankNOTAREALTYPE'
        self.validator.validate_predefined_type('IfcTank', 'NOTAREALTYPE', uri, 'Classes', 7)
        self.assertIn('invalid_predefined_type', self.finding_codes())

    def test_single_ifc_set_reference_is_accepted(self):
        references = self.validator.parse_ifc_linked_references('Pset_PipeSegmentTypeCommon', 'Properties', 60)
        self.assertEqual(['Pset_PipeSegmentTypeCommon'], references)
        self.assertEqual(set(), self.finding_codes())

    def test_multiple_ifc_set_references_require_json_array(self):
        references = self.validator.parse_ifc_linked_references('Pset_PipeSegmentTypeCommon\nQto_WallBaseQuantities', 'Properties', 60)
        self.assertEqual([], references)
        self.assertIn('invalid_ifc_linked_list_syntax', self.finding_codes())

    def test_json_ifc_set_reference_list_is_accepted(self):
        raw = '["Pset_PipeSegmentTypeCommon", "Qto_WallBaseQuantities"]'
        references = self.validator.parse_ifc_linked_references(raw, 'Properties', 60)
        self.assertEqual(['Pset_PipeSegmentTypeCommon', 'Qto_WallBaseQuantities'], references)
        self.assertEqual(set(), self.finding_codes())

    def test_ifc_set_reference_list_requires_non_empty_strings(self):
        references = self.validator.parse_ifc_linked_references('["Pset_WallCommon", ""]', 'Properties', 60)
        self.assertEqual([], references)
        self.assertIn('invalid_ifc_linked_list_syntax', self.finding_codes())

    def test_duplicate_ifc_set_reference_is_rejected(self):
        raw = '["Qto_WallBaseQuantities", "Qto_WallBaseQuantities"]'
        references = self.validator.parse_ifc_linked_references(raw, 'Properties', 60)
        self.assertEqual(['Qto_WallBaseQuantities', 'Qto_WallBaseQuantities'], references)
        self.assertIn('duplicate_ifc_linked_reference', self.finding_codes())


if __name__ == '__main__':
    unittest.main()
