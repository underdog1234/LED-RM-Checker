import datetime as dt
import importlib.util
import sys
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "check_rentman_led.py"
SPEC = importlib.util.spec_from_file_location("check_rentman_led", MODULE_PATH)
checker = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = checker
SPEC.loader.exec_module(checker)


class CheckRentmanLedTests(unittest.TestCase):
    def test_tag_is_present_handles_comma_separated_tags(self):
        self.assertTrue(checker.tag_is_present("auckland, tec✔️, led", "tec✔️"))
        self.assertFalse(checker.tag_is_present("auckland, led", "tec✔️"))

    def test_project_is_upcoming_uses_plan_period_end(self):
        timezone = dt.timezone.utc
        project = {
            "planperiod_start": "2026-07-14T09:00:00+00:00",
            "planperiod_end": "2026-07-16T17:00:00+00:00",
        }
        self.assertTrue(
            checker.project_is_upcoming(project, dt.date(2026, 7, 15), timezone, None)
        )
        self.assertFalse(
            checker.project_is_upcoming(project, dt.date(2026, 7, 17), timezone, None)
        )

    def test_exact_equipment_match_checks_equipment_code(self):
        target = checker.EquipmentTarget(
            row_number=2,
            values={"Code": "12223"},
            identifiers=frozenset({checker.normalize("12223")}),
        )
        item = {
            "id": 99,
            "name": "Different project label",
            "quantity": 4,
            "equipment": {"code": "12223", "name": "YES TECH panel"},
        }
        match = checker.match_equipment_item(item, [target], "exact")
        self.assertIsNotNone(match)
        self.assertEqual(match.target_row, 2)
        self.assertEqual(match.quantity, 4)

    def test_contains_match_can_match_longer_project_name(self):
        target = checker.EquipmentTarget(
            row_number=3,
            values={"Name (in database)": "YES TECH MG9 P2.9 500mm x 500mm LED Panel"},
            identifiers=frozenset(
                {checker.normalize("YES TECH MG9 P2.9 500mm x 500mm LED Panel")}
            ),
        )
        item = {
            "id": 100,
            "displayname": "YES TECH MG9 P2.9 500mm x 500mm LED Panel - kit",
            "quantity_total": 12,
            "equipment": {},
        }
        self.assertIsNone(checker.match_equipment_item(item, [target], "exact"))
        self.assertIsNotNone(checker.match_equipment_item(item, [target], "contains"))


if __name__ == "__main__":
    unittest.main()
