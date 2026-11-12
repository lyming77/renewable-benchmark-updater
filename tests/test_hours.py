"""Synthetic-only checks for the reusable utilization-hours engine."""

import itertools
import math
import sys
import unittest
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from benchmark.hours import (  # noqa: E402
    HoursFieldLimitError,
    HoursModelError,
    build_hours_model,
    evaluate_hours,
)


HEADERS = [f"字段{i + 1}" for i in range(67)]


def record(station, month, energy=100, capacity=10, region="甲地", company="示例企业", commissioned="2020-01-01", dispatch=None):
    values = [None] * 67
    values[0] = "合成数据"
    values[1] = f"{month}月份"
    values[2] = station
    values[3] = dispatch or station
    values[4] = region
    values[5] = "示例光伏"
    values[6] = capacity
    values[7] = commissioned
    values[20] = energy
    values[66] = company
    return {"values": values, "source": "synthetic.xlsx"}


def helper_evaluation(model, selected, filters=None):
    """Independent aggregation of emitted source buckets, without unit logic."""
    filters = filters or {}
    indices = {item["key"]: item["source_col_index"] for item in model["dimension_fields"]}
    totals = [0.0] * len(model["fields"])
    for values, meta in zip(model["rows"], model["row_metadata"]):
        if meta["month"] not in selected:
            continue
        if any(values[indices[key]] not in accepted for key, accepted in filters.items()):
            continue
        for index, value in enumerate(values):
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                totals[index] += value
    index = model["field_index"]
    months = [item for item in model["months"] if totals[index[item["flag_field"]]] > 0]
    days = sum(item["days"] for item in months)
    energy = cd = count = 0.0
    for bucket in model["stable_buckets"]:
        if all(totals[index[f"F{month:02d}"]] == 0 for month in bucket["missing_months"]):
            energy += totals[index[bucket["e_field"]]]
            cd += totals[index[bucket["cd_field"]]]
            count += totals[index[bucket["count_field"]]] / len(months) if months else 0
    for bucket in model["special_buckets"]:
        good_days = totals[index[bucket["good_days_field"]]]
        if good_days == days and good_days > 0:
            energy += totals[index[bucket["e_field"]]]
            cd += totals[index[bucket["cd_field"]]]
            count += 1
    return {"energy_mwh": energy, "capacity_days": cd, "station_count": count,
            "hours": energy * days / cd if cd > 0 else None}


class HoursTests(unittest.TestCase):
    def test_arbitrary_months_and_classification_filters_match_source_buckets(self):
        rows = [record("完整站", month, month * 100) for month in (1, 2, 3)]
        rows += [record("断月站", month, 500) for month in (1, 3)]
        rows += [record("分类变化站", month, 200, region="乙地" if month == 2 else "甲地") for month in (1, 2, 3)]
        rows += [record("真实零站", month, 0) for month in (1, 2, 3)]
        rows += [record("缺报数值站", month, None if month == 2 else 300) for month in (1, 2, 3)]
        model = build_hours_model(rows, HEADERS)
        for size in (1, 2, 3):
            for selected in itertools.combinations((1, 2, 3), size):
                for regions in (None, ["甲地"], ["乙地"], ["甲地", "乙地"], ["不存在"]):
                    filters = {"region": regions} if regions is not None else None
                    oracle = evaluate_hours(rows, selected, filters)
                    actual = helper_evaluation(model, selected, filters)
                    with self.subTest(selected=selected, regions=regions):
                        for key in ("energy_mwh", "capacity_days", "station_count"):
                            self.assertAlmostEqual(oracle[key], actual[key])
                        if oracle["hours"] is None:
                            self.assertIsNone(actual["hours"])
                        else:
                            self.assertAlmostEqual(oracle["hours"], actual["hours"])

    def test_missing_company_month_is_not_silently_ignored(self):
        rows = [record("只报一月", 1, company="企业甲"), record("只报二月", 2, company="企业乙")]
        model = build_hours_model(rows, HEADERS)
        result = helper_evaluation(model, [1, 2], {"co": ["企业甲"]})
        self.assertEqual(result["station_count"], 0)
        self.assertIsNone(result["hours"])
        self.assertEqual(model["summary_by_company"]["企业甲"]["HH_D"], 59)

    def test_capacity_is_weighted_by_calendar_days_and_not_used_as_identity(self):
        rows = [record("扩容站", 1, 310, 10), record("扩容站", 2, 280, 20)]
        result = evaluate_hours(rows, [1, 2])
        self.assertEqual(result["station_count"], 1)
        self.assertAlmostEqual(result["capacity_mw"], (10 * 31 + 20 * 28) / 59)
        self.assertAlmostEqual(result["hours"], 590 * 59 / (310 + 560))
        self.assertEqual(len(result["capacity_change_unit_ids"]), 1)

    def test_same_name_simultaneous_phases_split_only_on_explicit_dates(self):
        rows = [record("同名一期", month, 100, 10, commissioned="2020-01-01") for month in (1, 2)]
        rows += [record("同名一期", month, 300, 30, commissioned=date(2021, 1, 1)) for month in (1, 2)]
        model = build_hours_model(rows, HEADERS)
        self.assertEqual(len(model["units"]), 2)
        self.assertEqual(model["summary_all"]["station_count"], 2)
        self.assertEqual(model["summary_all"]["energy_mwh"], 800)
        ambiguous = [record("重复站", 1, 100, 10), record("重复站", 1, 200, 20)]
        with self.assertRaisesRegex(HoursModelError, "身份有歧义"):
            build_hours_model(ambiguous, HEADERS)

    def test_raw_cells_are_preserved_and_controls_do_not_pollute_numeric_metrics(self):
        rows = [record("保留站", 1)]
        rows[0]["values"][59] = 123.45
        before = rows[0]["values"].copy()
        model = build_hours_model(rows, HEADERS)
        self.assertEqual(model["rows"][0][:67], before)
        self.assertEqual(rows[0]["values"], before)
        for values, meta in zip(model["rows"], model["row_metadata"]):
            if meta["kind"] == "sentinel":
                for col in (0, 6, 20, 22, 59, 60):
                    self.assertIsNone(values[col])

    def test_true_zero_missing_text_bool_and_invalid_capacity_are_distinct(self):
        valid_zero = [record("零电量", 1, 0)]
        self.assertEqual(evaluate_hours(valid_zero, [1])["hours"], 0)
        for energy, capacity in ((None, 10), ("0", 10), (True, 10), (math.nan, 10), (100, 0), (100, -1), (100, math.inf)):
            with self.subTest(energy=energy, capacity=capacity):
                result = evaluate_hours([record("无效站", 1, energy, capacity)], [1])
                self.assertIsNone(result["hours"])
                self.assertEqual(result["station_count"], 0)

    def test_twelve_months_and_leap_year_are_dynamic(self):
        rows = [record("年度站", month, 100) for month in range(1, 13)]
        model = build_hours_model(rows, HEADERS, year=2028)
        self.assertEqual(len(model["months"]), 12)
        self.assertEqual(model["months"][1]["days"], 29)
        self.assertEqual(model["summary_all"]["calendar_days"], 366)
        self.assertAlmostEqual(model["summary_all"]["hours"], 120)

    def test_exact_blank_classification_and_multi_selection(self):
        rows = [record("分类缺报站", 1, region=None), record("分类缺报站", 2, region="甲地")]
        self.assertEqual(evaluate_hours(rows, [1, 2])["station_count"], 1)
        self.assertEqual(evaluate_hours(rows, [1, 2], {"region": "甲地"})["station_count"], 0)
        self.assertEqual(evaluate_hours(rows, [1, 2], {"region": [None, "甲地"]})["station_count"], 1)

    def test_field_limit_is_explicit_instead_of_dropping_stations(self):
        rows = [record(f"变动站{station}", month, region=f"地区{month}") for station in range(55) for month in (1, 2)]
        with self.assertRaisesRegex(HoursFieldLimitError, "超过WPS普通透视表255"):
            build_hours_model(rows, HEADERS)

    def test_native_formulas_avoid_dynamic_and_functions(self):
        rows = [record("稳定站", 1), record("跨类站", 1, region="甲地"),
                record("跨类站", 2, region="乙地")]
        model = build_hours_model(rows, HEADERS)
        for formula in [item["formula"] for item in model["formulas"]] + [model["hours_formula"]]:
            self.assertNotIn("AND(", formula.upper())
            self.assertNotIn("OR(", formula.upper())
        self.assertEqual(model["metadata"]["formula_functions"], ["IF"])
        formulas = {item["name"]: item["formula"] for item in model["formulas"]}
        # Exact native shapes follow the previously WPS-verified IF-only model.
        # These constants contain synthetic bucket IDs and no business data.
        for name in ("HH_E", "HH_CD", "HH_Count", "HH_C"):
            self.assertTrue(formulas[name].startswith("IF('小时源变更'=0,IF(HH_D>0,"))
            self.assertTrue(formulas[name].endswith(",0),0)"))
        self.assertEqual(formulas["HH_E"],
                         "IF('小时源变更'=0,IF(HH_D>0,IF('F02'=0,E_P01,0)"
                         "+IF(GD_U01=HH_D,E_U01,0),0),0)")
        self.assertEqual(model["hours_formula"],
                         "IF('小时源变更'=0,IF(HH_D>0,IF(HH_CD>0,"
                         "HH_E*HH_D/HH_CD,0/0),0/0),0/0)")

    def test_month_labels_year_and_filter_names_are_validated(self):
        rows = [record("站点", 1), record("站点", 2)]
        rows[1]["values"][1] = "2027年2月"
        with self.assertRaisesRegex(HoursModelError, "不属于统计年度"):
            build_hours_model(rows, HEADERS)
        with self.assertRaisesRegex(HoursModelError, "不支持筛选字段"):
            evaluate_hours([record("站点", 1)], [1], {"capacity": [10]})


if __name__ == "__main__":
    unittest.main()
