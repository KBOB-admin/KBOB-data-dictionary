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


if __name__ == '__main__':
    unittest.main()
