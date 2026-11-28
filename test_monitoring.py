from pathlib import Path
import pandas as pd
import yaml

from reserve_monitor import normalize_headers, prepare_snapshot, build_movements, daily_summary


def test_bridge(sample_prev: str, sample_curr: str, config_path: str = "config.yaml"):
    cfg = yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))
    p = pd.read_csv(sample_prev, sep=None, engine="python")
    c = pd.read_csv(sample_curr, sep=None, engine="python")
    prev, _ = prepare_snapshot(p, cfg)
    curr, report_date = prepare_snapshot(c, cfg)
    movements = build_movements(curr, prev, cfg)
    summary = daily_summary(curr, movements, report_date).iloc[0]
    assert abs(float(summary["bridge_check_mln"])) < 1e-9
    expected = curr["reserve_mln"].sum() - prev["reserve_mln"].sum()
    assert abs(float(summary["delta_reserve_dod_mln"]) - float(expected)) < 1e-9
    print("OK", report_date.date(), summary["delta_reserve_dod_mln"])


if __name__ == "__main__":
    test_bridge(
        "../sample/sample_raw_2026-09-27.csv",
        "../sample/sample_raw_2026-09-28.csv",
    )
