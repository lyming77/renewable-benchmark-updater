"""Purely synthetic preview checks; no uploaded or historical business data."""

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from benchmark.service import preview_summary  # noqa: E402


def row(station, month, *, company="样例甲", station_type="光伏", energy=100,
        capacity=10, settled=100, realtime_fee=10000,
        commissioned="2020-01-01"):
    values = [None] * 67
    values[1] = f"{month}月份"
    values[2] = station
    values[3] = station
    values[4] = "样例地区"
    values[5] = station_type
    values[6] = capacity
    values[7] = commissioned
    values[20] = energy
    values[22] = realtime_fee
    values[59] = settled
    values[66] = company
    return {"values": values, "source": "synthetic.xlsx", "month": month,
            "year": 2026, "company": company}


def aggregate(result):
    return next(item for item in result if item["集团"] == "全部所选集团")


def test_preview_noncontiguous_months_preserve_original_phase_identities():
    # February proves these are two independent same-name phases.  Choosing
    # January + March must not erase that identity evidence and merge phases.
    records = [row("同名分期", month, commissioned="2020-01-01") for month in (1, 2)]
    records += [row("同名分期", month, commissioned="2021-01-01") for month in (2, 3)]
    result = aggregate(preview_summary(records, [1, 3], ["样例甲"], ["光伏"], 2026))
    assert result["完整样本站数"] == 0
    assert result["累计利用小时（h）"] is None
    assert result["总结算电量（MWh）"] == 200


def test_preview_type_filter_cannot_merge_different_same_name_phases():
    # A second type supplies the only simultaneous observation.  Filtering it
    # out must not change the independent identity registry.
    records = [row("同名分期", 1, commissioned="2020-01-01"),
               row("同名分期", 2, station_type="风电", commissioned="2020-01-01"),
               row("同名分期", 2, commissioned="2021-01-01"),
               row("同名分期", 3, commissioned="2021-01-01")]
    result = aggregate(preview_summary(records, [1, 2, 3], ["样例甲"], ["光伏"], 2026))
    assert result["完整样本站数"] == 0
    assert result["累计利用小时（h）"] is None
    assert result["总结算电量（MWh）"] == 300


def test_trade_population_stays_full_while_hours_use_complete_population():
    records = [row("完整站", month, energy=100, settled=100, realtime_fee=10000) for month in (1, 2)]
    records += [row("短报站", 1, energy=1000, settled=1000, realtime_fee=200000)]
    result = aggregate(preview_summary(records, [1, 2], ["样例甲"], ["光伏"], 2026))
    assert result["完整样本站数"] == 1
    assert result["累计利用小时（h）"] == pytest.approx(20)
    assert result["总结算电量（MWh）"] == 1200
    assert result["实时电价贡献（元/MWh）"] == pytest.approx(220000 / 1200)
    assert result["交易电价（元/MWh）"] == pytest.approx(220000 / 1200)


def test_shared_company_and_type_selections_apply_to_both_populations():
    records = [row("甲光伏", month, energy=100) for month in (1, 2)]
    records += [row("甲风电", month, station_type="风电", energy=300) for month in (1, 2)]
    records += [row("乙光伏", month, company="样例乙", energy=500) for month in (1, 2)]
    result = aggregate(preview_summary(records, [1, 2], ["样例甲"], ["光伏"], 2026))
    assert result["完整样本站数"] == 1
    assert result["累计利用小时（h）"] == pytest.approx(20)
    assert result["总结算电量（MWh）"] == 200
    combined = aggregate(preview_summary(records, [1, 2], ["样例甲", "样例乙"], ["光伏"], 2026))
    assert combined["完整样本站数"] == 2
    assert combined["累计利用小时（h）"] == pytest.approx(60)
    assert combined["总结算电量（MWh）"] == 400


def test_type_union_restores_full_history_of_changing_station():
    records = [row("跨类站", 1, station_type="光伏"), row("跨类站", 2, station_type="风电")]
    restricted = aggregate(preview_summary(records, [1, 2], ["样例甲"], ["光伏"], 2026))
    assert restricted["完整样本站数"] == 0
    restored = aggregate(preview_summary(records, [1, 2], ["样例甲"], ["光伏", "风电"], 2026))
    assert restored["完整样本站数"] == 1
    assert restored["累计利用小时（h）"] == pytest.approx(20)


def test_zero_energy_is_valid_and_empty_selection_is_empty():
    records = [row("真实零站", month, energy=0) for month in (1, 2)]
    result = aggregate(preview_summary(records, [1, 2], ["样例甲"], ["光伏"], 2026))
    assert result["完整样本站数"] == 1
    assert result["累计利用小时（h）"] == 0
    assert preview_summary(records, [], ["样例甲"], ["光伏"], 2026) == []
