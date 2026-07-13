import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "generate_booking_data.py"
SPEC = importlib.util.spec_from_file_location("generate_booking_data", MODULE_PATH)
GENERATOR = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(GENERATOR)


def booking_records(pol, pod, lane, vvd):
    record = [None] * 13
    record[12] = [[pol, pod, lane, vvd]]
    return [record]


class RouteOverrideTests(unittest.TestCase):
    def test_prefers_earlier_cross_lane_physical_vessel_rotation(self):
        schedule = {
            "voyages": {
                "MV1|SNPO2635S": [
                    [
                        "VNSGN",
                        "2026-07-07 17:30",
                        "2026-07-07 20:55",
                        "2026-07-08 16:10",
                        1,
                        100,
                    ]
                ],
                "MI1|SNPO2627S": [
                    [
                        "MYPKG",
                        "2026-07-10 23:00",
                        "2026-07-11 00:35",
                        "2026-07-12 00:45",
                        1,
                        200,
                    ]
                ],
                "MV1|SNPO2629N": [
                    [
                        "MYPKG",
                        "2026-07-17 06:00",
                        "2026-07-17 08:00",
                        "2026-07-18 06:00",
                        1,
                        300,
                    ]
                ],
            }
        }

        result = GENERATOR.build_route_overrides(
            booking_records("VNSGN", "MYPKG", "MV1", "SNPO2635S"),
            schedule,
        )

        selected = result["overrides"]["MV1|SNPO2635S|VNSGN|MYPKG"][1][0]
        self.assertEqual(selected[0], "SNPO2627S")
        self.assertEqual(selected[2], "2026-07-11 00:35")
        self.assertEqual(selected[6], "MI1")

    def test_keeps_same_lane_opposite_direction_when_it_is_next_rotation(self):
        schedule = {
            "voyages": {
                "CGX|XPBL2625W": [
                    [
                        "CNNGB",
                        "2026-06-01 08:00",
                        "2026-06-01 10:00",
                        "2026-06-01 12:00",
                        1,
                        10,
                    ]
                ],
                "CGX|XPBL2625E": [
                    [
                        "PKKHI",
                        "2026-06-08 08:00",
                        "2026-06-08 10:00",
                        "2026-06-08 12:00",
                        1,
                        20,
                    ]
                ],
            }
        }

        result = GENERATOR.build_route_overrides(
            booking_records("CNNGB", "PKKHI", "CGX", "XPBL2625W"),
            schedule,
        )

        selected = result["overrides"]["CGX|XPBL2625W|CNNGB|PKKHI"][1][0]
        self.assertEqual(selected[0], "XPBL2625E")
        self.assertEqual(selected[6], "CGX")

    def test_does_not_override_an_ordered_exact_voyage_path(self):
        schedule = {
            "voyages": {
                "SEA|ABCD2601S": [
                    [
                        "VNSGN",
                        "2026-06-01 08:00",
                        "2026-06-01 10:00",
                        "2026-06-01 12:00",
                        1,
                        10,
                    ],
                    [
                        "MYPKG",
                        "2026-06-03 08:00",
                        "2026-06-03 10:00",
                        "2026-06-03 12:00",
                        2,
                        20,
                    ],
                ]
            }
        }

        result = GENERATOR.build_route_overrides(
            booking_records("VNSGN", "MYPKG", "SEA", "ABCD2601S"),
            schedule,
        )

        self.assertNotIn("SEA|ABCD2601S|VNSGN|MYPKG", result["overrides"])


if __name__ == "__main__":
    unittest.main()
