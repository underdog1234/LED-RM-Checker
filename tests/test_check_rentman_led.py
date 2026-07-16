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

    def test_parse_datetime_ignores_timezone_overflow(self):
        timezone = dt.timezone(dt.timedelta(hours=13))
        self.assertIsNone(checker.parse_datetime("9999-12-31T23:59:59+00:00", timezone))

    def test_cancelled_project_is_identified_by_status_path(self):
        self.assertTrue(checker.project_is_cancelled({"status": "/statuses/2"}))
        self.assertFalse(checker.project_is_cancelled({"status": "/statuses/3"}))

    def test_project_status_can_come_from_subproject(self):
        project = {"_subprojects": [{"status": {"id": 4, "name": "Prepped"}}]}
        self.assertEqual(checker.project_status_details(project), ("📦", "Prepped"))

    def test_project_label_links_project_number_to_rentman(self):
        project = {"id": 4549, "number": 7042, "name": "Conference"}
        self.assertEqual(
            checker.project_label(project),
            "<https://multimedia.rentmanapp.com/#/projects/4549/details|#7042> Conference",
        )

    def test_project_date_summary_is_human_readable(self):
        project = {
            "planperiod_start": "2026-10-17T00:00:00+13:00",
            "planperiod_end": "2026-10-18T23:59:00+13:00",
        }
        timezone = dt.timezone(dt.timedelta(hours=13))
        self.assertEqual(
            checker.project_date_summary(project, timezone),
            "17/10/2026 00:00 to 18/10/2026 23:59",
        )

    def test_slack_message_reports_status_and_project_type(self):
        finding = checker.Finding(
            project={
                "id": 4549,
                "number": 7042,
                "name": "Conference",
                "status": "/statuses/3",
                "project_type": {"name": "Conference", "color": "ec00a3"},
                "planperiod_start": "2026-10-17T00:00:00+13:00",
                "planperiod_end": "2026-10-18T23:59:00+13:00",
            },
            matches=(
                checker.EquipmentMatch(
                    project_equipment_id=1,
                    display_name="YES TECH MG9 P2.9 LED Panel",
                    quantity=96,
                    matched_identifier="yes tech",
                    target_row=3,
                ),
            ),
        )
        message = checker.format_slack_message(
            [finding],
            "tec✔️",
            dt.timezone(dt.timedelta(hours=13)),
        )
        self.assertIn("Status: 👍 Confirmed", message)
        self.assertIn("Type: ", message)
        self.assertIn("Conference", message)
        self.assertIn("Planning period: 17/10/2026 00:00 to 18/10/2026 23:59", message)

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
