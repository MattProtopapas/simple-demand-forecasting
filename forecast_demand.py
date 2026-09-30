"""Forecast next month's demand per SKU and write everything to one Excel file.

Reads the validated CSVs produced by validate_data.py, holds back the last
HOLDOUT_WEEKS of history, fits three simple forecasting methods on the rest,
scores them on the holdout, picks a winner per SKU, and re-fits the winner on
the full history to forecast the next FORECAST_HORIZON_WEEKS. Writes a
brand-new workbook (no template) with these sheets:

  - Input Data            the validated weekly sales the forecasts ran on
  - SKU Master            static SKU attributes (price, lead time, service level)
  - Accuracy League Table one row per SKU: MAE / MAPE / bias per method, the
                          winning method and a forecastability class
  - Forecast vs Actual    holdout weeks per SKU, every method's forecast next
                          to the actual, with the winner's error
  - Next Month Forecast   the winner's forecast for the coming weeks, and the
                          safety stock / reorder point it implies
  - Forecastability       roll-up by class (Easy / Moderate / Hard) and a count
                          of which method won how often
  - Charts                forecast vs actual line charts for the highest-volume SKUs

Methods (all pure Python, no numpy / statsmodels):
  - Seasonal naive:   forecast = the value from the same week one season ago
  - Moving average:   forecast = mean of the last MOVING_AVERAGE_WINDOW weeks (flat)
  - Holt-Winters:     additive level + trend + seasonality (ETS(A,A,A)), with
                      the smoothing parameters chosen by grid search on
                      in-sample one-step-ahead squared error

Accuracy on the holdout:
  - MAE   = mean |forecast - actual|
  - MAPE  = mean |forecast - actual| / actual   (weeks with actual = 0 skipped)
  - RMSE  = sqrt(mean (forecast - actual)^2)
  - Bias  = mean (forecast - actual)   (positive = over-forecasting)

Link to inventory:
  - Safety stock = z(service level) * RMSE of the winning forecast * sqrt(lead time in weeks)
  - Reorder point = forecast demand over the lead time + safety stock
  For comparison the sheet also shows the safety stock you would hold if you
  ignored the forecast and used the raw std. dev. of weekly demand instead.
"""

import argparse
import csv
import math
import os
from collections import defaultdict
from datetime import date, timedelta
from statistics import NormalDist, mean, pstdev

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from forecast_common import (
    FORECAST_HORIZON_WEEKS,
    FORECASTABILITY_BANDS,
    HOLDOUT_WEEKS,
    MOVING_AVERAGE_WINDOW,
    OUTPUT_WORKBOOK_PATH,
    SALES_FIELDS,
    SALES_VALIDATED_PATH,
    SEASON_LENGTH,
    SKU_FIELDS,
    SKU_MASTER_VALIDATED_PATH,
    SalesRecord,
    SkuRecord,
)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)
HARD_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
MODERATE_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

METHODS = ["Seasonal Naive", "Moving Average", "Holt-Winters"]

# Chart colours: actual in a neutral ink, the methods in the first three
# categorical slots (blue, orange, aqua).
ACTUAL_COLOR = "52514E"
METHOD_COLORS = {"Seasonal Naive": "2A78D6", "Moving Average": "EB6834", "Holt-Winters": "1BAF7A"}
CHART_SKU_COUNT = 6

LEAGUE_FIELDS = [
    "rank",
    "sku",
    "item_name",
    "category",
    "avg_weekly_demand",
    "demand_cv",
    "pct_zero_weeks",
    "weeks_filled",
    "mae_seasonal_naive",
    "mae_moving_average",
    "mae_holt_winters",
    "mape_seasonal_naive",
    "mape_moving_average",
    "mape_holt_winters",
    "bias_seasonal_naive",
    "bias_moving_average",
    "bias_holt_winters",
    "best_method",
    "best_mae",
    "best_mape",
    "best_bias_pct",
    "forecastability",
]

FVA_FIELDS = [
    "sku",
    "week_start",
    "on_promo",
    "actual",
    "seasonal_naive",
    "moving_average",
    "holt_winters",
    "best_method",
    "best_forecast",
    "error",
    "abs_pct_error",
]

NEXT_MONTH_FIELDS = [
    "sku",
    "item_name",
    "category",
    "forecastability",
    "best_method",
    *[f"week_{i + 1}" for i in range(FORECAST_HORIZON_WEEKS)],
    "next_month_total",
    "avg_per_week",
    "lead_time_weeks",
    "service_level",
    "forecast_rmse",
    "safety_stock_from_forecast_error",
    "safety_stock_from_demand_std",
    "lead_time_demand",
    "reorder_point",
]

SUMMARY_FIELDS = [
    "forecastability",
    "mape_range",
    "sku_count",
    "pct_of_skus",
    "share_of_units",
    "avg_best_mape",
    "won_by_seasonal_naive",
    "won_by_moving_average",
    "won_by_holt_winters",
]


# ---------------------------------------------------------------------------
# Loading and reshaping
# ---------------------------------------------------------------------------


def read_skus(path: str) -> dict[str, SkuRecord]:
    with open(path, newline="", encoding="utf-8") as f:
        return {row["sku"]: SkuRecord(**row) for row in csv.DictReader(f)}


def read_sales(path: str) -> list[SalesRecord]:
    with open(path, newline="", encoding="utf-8") as f:
        return [SalesRecord(**row) for row in csv.DictReader(f)]


def build_series(sales: list[SalesRecord]) -> tuple[list[date], dict[str, list[float]], dict[str, int]]:
    """Turn the long sales table into one gap-free weekly series per SKU.

    Weeks dropped by the validator are filled by linear interpolation between
    their neighbours (or copied from the nearest week at the edges) so every
    SKU has a value for every week. The number of filled weeks is reported per
    SKU so nobody mistakes an interpolated value for a real one.
    """
    by_sku: dict[str, dict[date, float]] = defaultdict(dict)
    for rec in sales:
        by_sku[rec.sku][rec.week_start] = float(rec.units_sold)

    first = min(rec.week_start for rec in sales)
    last = max(rec.week_start for rec in sales)
    weeks = [first + timedelta(weeks=i) for i in range((last - first).days // 7 + 1)]

    series: dict[str, list[float]] = {}
    filled: dict[str, int] = {}
    for sku, values in by_sku.items():
        raw = [values.get(w) for w in weeks]
        series[sku] = interpolate_gaps(raw)
        filled[sku] = sum(1 for v in raw if v is None)
    return weeks, series, filled


def interpolate_gaps(values: list[float | None]) -> list[float]:
    known = [i for i, v in enumerate(values) if v is not None]
    if not known:
        return [0.0] * len(values)
    out: list[float] = []
    for i, v in enumerate(values):
        if v is not None:
            out.append(v)
            continue
        before = [k for k in known if k < i]
        after = [k for k in known if k > i]
        if before and after:
            b, a = before[-1], after[0]
            frac = (i - b) / (a - b)
            out.append(values[b] + frac * (values[a] - values[b]))
        elif before:
            out.append(values[before[-1]])
        else:
            out.append(values[after[0]])
    return out


# ---------------------------------------------------------------------------
# Forecasting methods
# ---------------------------------------------------------------------------


def seasonal_naive(history: list[float], horizon: int) -> list[float]:
    n = len(history)
    if n < SEASON_LENGTH:
        return [history[-1]] * horizon
    return [history[n - SEASON_LENGTH + (h % SEASON_LENGTH)] for h in range(horizon)]


def moving_average(history: list[float], horizon: int) -> list[float]:
    window = history[-MOVING_AVERAGE_WINDOW:]
    return [mean(window)] * horizon


def holt_winters_run(
    history: list[float], alpha: float, beta: float, gamma: float, m: int
) -> tuple[float, float, list[float], float]:
    """Run additive Holt-Winters over the history.

    Returns the final level, trend, seasonal indices and the in-sample
    one-step-ahead SSE (measured after the first season used for
    initialisation).
    """
    level = mean(history[:m])
    trend = (mean(history[m : 2 * m]) - level) / m
    season = [history[i] - level for i in range(m)]

    sse = 0.0
    for t, y in enumerate(history):
        s = season[t % m]
        if t >= m:
            sse += (level + trend + s - y) ** 2
        new_level = alpha * (y - s) + (1 - alpha) * (level + trend)
        trend = beta * (new_level - level) + (1 - beta) * trend
        season[t % m] = gamma * (y - new_level) + (1 - gamma) * s
        level = new_level
    return level, trend, season, sse


def holt_winters(history: list[float], horizon: int) -> list[float]:
    m = SEASON_LENGTH
    if len(history) < 2 * m:
        # not enough history to initialise a seasonal model; fall back
        return moving_average(history, horizon)

    best = None
    for alpha in (0.05, 0.1, 0.2, 0.3, 0.5):
        for beta in (0.01, 0.05, 0.1, 0.2):
            for gamma in (0.05, 0.1, 0.2, 0.3, 0.5):
                level, trend, season, sse = holt_winters_run(history, alpha, beta, gamma, m)
                if best is None or sse < best[0]:
                    best = (sse, level, trend, season)

    _, level, trend, season = best
    n = len(history)
    return [max(0.0, level + (h + 1) * trend + season[(n + h) % m]) for h in range(horizon)]


FORECASTERS = {
    "Seasonal Naive": seasonal_naive,
    "Moving Average": moving_average,
    "Holt-Winters": holt_winters,
}


# ---------------------------------------------------------------------------
# Scoring
# ---------------------------------------------------------------------------


def score(forecast: list[float], actual: list[float]) -> dict:
    errors = [f - a for f, a in zip(forecast, actual)]
    abs_errors = [abs(e) for e in errors]
    pct_errors = [abs(f - a) / a for f, a in zip(forecast, actual) if a > 0]
    avg_actual = mean(actual) if actual else 0.0
    bias = mean(errors)
    return {
        "mae": mean(abs_errors),
        "rmse": math.sqrt(mean(e**2 for e in errors)),
        "mape": mean(pct_errors) if pct_errors else None,
        "bias": bias,
        "bias_pct": bias / avg_actual if avg_actual else None,
    }


def classify_forecastability(mape: float | None) -> str:
    if mape is None:
        return FORECASTABILITY_BANDS[-1][0]
    for label, upper in FORECASTABILITY_BANDS:
        if mape <= upper:
            return label
    return FORECASTABILITY_BANDS[-1][0]


def slug(method: str) -> str:
    return method.lower().replace("-", "_").replace(" ", "_")


def pct(value: float | None, digits: int = 1) -> float | None:
    return None if value is None else round(value * 100, digits)


# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------


def analyze(
    skus: dict[str, SkuRecord],
    weeks: list[date],
    series: dict[str, list[float]],
    filled: dict[str, int],
    promo_lookup: dict[tuple[str, date], bool],
) -> tuple[list[dict], list[dict], list[dict]]:
    league: list[dict] = []
    fva_rows: list[dict] = []
    next_month: list[dict] = []

    holdout_weeks = weeks[-HOLDOUT_WEEKS:]

    for sku, history in sorted(series.items()):
        record = skus[sku]
        train, holdout = history[:-HOLDOUT_WEEKS], history[-HOLDOUT_WEEKS:]

        forecasts = {name: fn(train, HOLDOUT_WEEKS) for name, fn in FORECASTERS.items()}
        scores = {name: score(fc, holdout) for name, fc in forecasts.items()}

        best_method = min(METHODS, key=lambda mth: scores[mth]["mae"])
        best = scores[best_method]
        forecastability = classify_forecastability(best["mape"])

        last_year = history[-SEASON_LENGTH:]
        avg_demand = mean(last_year)
        row = {
            "sku": sku,
            "item_name": record.item_name,
            "category": record.category,
            "avg_weekly_demand": round(avg_demand, 1),
            "demand_cv": round(pstdev(last_year) / avg_demand, 2) if avg_demand else None,
            "pct_zero_weeks": pct(sum(1 for v in history if v == 0) / len(history)),
            "weeks_filled": filled[sku],
            "best_method": best_method,
            "best_mae": round(best["mae"], 1),
            "best_mape": pct(best["mape"]),
            "best_bias_pct": pct(best["bias_pct"]),
            "forecastability": forecastability,
        }
        for name in METHODS:
            row[f"mae_{slug(name)}"] = round(scores[name]["mae"], 1)
            row[f"mape_{slug(name)}"] = pct(scores[name]["mape"])
            row[f"bias_{slug(name)}"] = round(scores[name]["bias"], 1)
        league.append(row)

        for i, week in enumerate(holdout_weeks):
            actual = holdout[i]
            best_fc = forecasts[best_method][i]
            fva_rows.append(
                {
                    "sku": sku,
                    "week_start": week,
                    "on_promo": promo_lookup.get((sku, week), False),
                    "actual": round(actual),
                    **{slug(name): round(forecasts[name][i], 1) for name in METHODS},
                    "best_method": best_method,
                    "best_forecast": round(best_fc, 1),
                    "error": round(best_fc - actual, 1),
                    "abs_pct_error": pct(abs(best_fc - actual) / actual) if actual else None,
                }
            )

        # Re-fit the winner on the full history for the real forecast.
        horizon = max(FORECAST_HORIZON_WEEKS, record.lead_time_weeks)
        future = FORECASTERS[best_method](history, horizon)
        next_4 = future[:FORECAST_HORIZON_WEEKS]

        z = NormalDist().inv_cdf(record.service_level)
        lt = record.lead_time_weeks
        ss_forecast = z * best["rmse"] * math.sqrt(lt)
        ss_demand = z * pstdev(last_year) * math.sqrt(lt)
        lead_time_demand = sum(future[:lt])

        next_month.append(
            {
                "sku": sku,
                "item_name": record.item_name,
                "category": record.category,
                "forecastability": forecastability,
                "best_method": best_method,
                **{f"week_{i + 1}": round(v) for i, v in enumerate(next_4)},
                "next_month_total": round(sum(next_4)),
                "avg_per_week": round(mean(next_4), 1),
                "lead_time_weeks": lt,
                "service_level": record.service_level,
                "forecast_rmse": round(best["rmse"], 1),
                "safety_stock_from_forecast_error": round(ss_forecast),
                "safety_stock_from_demand_std": round(ss_demand),
                "lead_time_demand": round(lead_time_demand),
                "reorder_point": round(lead_time_demand + ss_forecast),
            }
        )

    # Rank the league table by the winner's MAPE (unscorable SKUs go last).
    league.sort(key=lambda r: (r["best_mape"] is None, r["best_mape"] if r["best_mape"] is not None else 0))
    for i, row in enumerate(league, start=1):
        row["rank"] = i

    return league, fva_rows, next_month


def build_summary(league: list[dict]) -> list[dict]:
    total_skus = len(league) or 1
    total_units = sum(r["avg_weekly_demand"] for r in league) or 1.0
    summary = []
    lower = 0.0
    for label, upper in FORECASTABILITY_BANDS:
        rows = [r for r in league if r["forecastability"] == label]
        mapes = [r["best_mape"] for r in rows if r["best_mape"] is not None]
        if math.isinf(upper):
            mape_range = f"> {round(lower * 100)}%"
        else:
            mape_range = f"{round(lower * 100)}% - {round(upper * 100)}%"
        summary.append(
            {
                "forecastability": label,
                "mape_range": mape_range,
                "sku_count": len(rows),
                "pct_of_skus": round(len(rows) / total_skus * 100, 1),
                "share_of_units": round(sum(r["avg_weekly_demand"] for r in rows) / total_units * 100, 1),
                "avg_best_mape": round(mean(mapes), 1) if mapes else None,
                **{f"won_by_{slug(m)}": sum(1 for r in rows if r["best_method"] == m) for m in METHODS},
            }
        )
        lower = upper
    return summary


# ---------------------------------------------------------------------------
# Excel output
# ---------------------------------------------------------------------------


def autosize_columns(ws) -> None:
    for col_cells in ws.columns:
        length = max((len(str(c.value)) for c in col_cells if c.value is not None), default=10)
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(length + 2, 40)


def write_sheet(ws, fieldnames: list[str], rows: list[dict]) -> None:
    ws.append(fieldnames)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append([row.get(f, "") for f in fieldnames])
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(fieldnames))}{len(rows) + 1}"
    autosize_columns(ws)


def highlight_forecastability(ws, fieldnames: list[str]) -> None:
    col = fieldnames.index("forecastability")
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row):
        label = row_cells[col].value
        fill = HARD_FILL if label == "Hard" else MODERATE_FILL if label == "Moderate" else None
        if fill:
            for cell in row_cells:
                cell.fill = fill


def style_line(series, color: str, width_emu: int) -> None:
    series.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=color, w=width_emu))
    series.smooth = False
    series.marker.symbol = "none"


def add_charts(ws_charts, ws_fva, fva_rows: list[dict], league: list[dict]) -> None:
    """One forecast-vs-actual line chart per high-volume SKU, reading the
    Forecast vs Actual sheet directly so the charts stay live in Excel."""
    top = sorted(league, key=lambda r: r["avg_weekly_demand"], reverse=True)[:CHART_SKU_COUNT]
    name_by_sku = {r["sku"]: r["item_name"] for r in league}
    mape_by_sku = {r["sku"]: r["best_mape"] for r in league}

    col = {name: FVA_FIELDS.index(name) + 1 for name in FVA_FIELDS}
    anchor_row = 1
    for row in top:
        sku = row["sku"]
        indices = [i for i, r in enumerate(fva_rows) if r["sku"] == sku]
        if not indices:
            continue
        first, last = indices[0] + 2, indices[-1] + 2  # +2: header row and 1-based

        chart = LineChart()
        mape = mape_by_sku[sku]
        mape_text = f"{mape}% MAPE" if mape is not None else "MAPE n/a"
        chart.title = f"{sku} {name_by_sku[sku]} - holdout weeks ({row['best_method']} wins, {mape_text})"
        chart.y_axis.title = "Units / week"
        chart.x_axis.number_format = "dd mmm"
        chart.height, chart.width = 8, 18
        chart.legend.position = "b"

        for field, color, width in (
            ("actual", ACTUAL_COLOR, 28575),
            ("seasonal_naive", METHOD_COLORS["Seasonal Naive"], 19050),
            ("moving_average", METHOD_COLORS["Moving Average"], 19050),
            ("holt_winters", METHOD_COLORS["Holt-Winters"], 19050),
        ):
            ref = Reference(ws_fva, min_col=col[field], min_row=first, max_row=last)
            chart.add_data(ref, titles_from_data=False)
            series = chart.series[-1]
            series.tx = SeriesLabel(v=field.replace("_", " ").title())
            style_line(series, color, width)

        chart.set_categories(Reference(ws_fva, min_col=col["week_start"], min_row=first, max_row=last))

        ws_charts.add_chart(chart, f"A{anchor_row}")
        anchor_row += 17

    ws_charts["T1"] = "Charts show the last quarter of history (the holdout) that every method was scored on."
    ws_charts["T2"] = "Actual is the dark line; each method is one colour throughout the workbook."


def build_workbook(
    input_rows: list[dict],
    sku_rows: list[dict],
    league: list[dict],
    fva_rows: list[dict],
    next_month: list[dict],
    summary: list[dict],
) -> Workbook:
    wb = Workbook()

    ws_input = wb.active
    ws_input.title = "Input Data"
    write_sheet(ws_input, SALES_FIELDS, input_rows)

    ws_skus = wb.create_sheet("SKU Master")
    write_sheet(ws_skus, SKU_FIELDS, sku_rows)

    ws_league = wb.create_sheet("Accuracy League Table")
    write_sheet(ws_league, LEAGUE_FIELDS, league)
    highlight_forecastability(ws_league, LEAGUE_FIELDS)

    ws_fva = wb.create_sheet("Forecast vs Actual")
    write_sheet(ws_fva, FVA_FIELDS, fva_rows)
    for cell in ws_fva["B"][1:]:
        cell.number_format = "yyyy-mm-dd"

    ws_next = wb.create_sheet("Next Month Forecast")
    write_sheet(ws_next, NEXT_MONTH_FIELDS, next_month)
    highlight_forecastability(ws_next, NEXT_MONTH_FIELDS)

    ws_summary = wb.create_sheet("Forecastability")
    write_sheet(ws_summary, SUMMARY_FIELDS, summary)

    ws_charts = wb.create_sheet("Charts")
    add_charts(ws_charts, ws_fva, fva_rows, league)

    return wb


def main() -> None:
    parser = argparse.ArgumentParser(description="Forecast weekly demand per SKU and write results to Excel.")
    parser.add_argument("--sales-input", type=str, default=SALES_VALIDATED_PATH, help="Path to the validated sales CSV")
    parser.add_argument(
        "--master-input", type=str, default=SKU_MASTER_VALIDATED_PATH, help="Path to the validated SKU master CSV"
    )
    parser.add_argument("--output", type=str, default=OUTPUT_WORKBOOK_PATH, help="Path to write the Excel workbook to")
    args = parser.parse_args()

    skus = read_skus(args.master_input)
    sales = read_sales(args.sales_input)
    with open(args.sales_input, newline="", encoding="utf-8") as f:
        input_rows = list(csv.DictReader(f))
    with open(args.master_input, newline="", encoding="utf-8") as f:
        sku_rows = list(csv.DictReader(f))

    weeks, series, filled = build_series(sales)
    if len(weeks) <= HOLDOUT_WEEKS:
        raise SystemExit(f"Need more than {HOLDOUT_WEEKS} weeks of history to hold out {HOLDOUT_WEEKS}.")
    promo_lookup = {(rec.sku, rec.week_start): rec.on_promo for rec in sales}

    league, fva_rows, next_month = analyze(skus, weeks, series, filled, promo_lookup)
    summary = build_summary(league)

    wb = build_workbook(input_rows, sku_rows, league, fva_rows, next_month, summary)
    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    wb.save(args.output)

    print(f"Forecast {len(league)} SKUs over {len(weeks)} weeks of history -> {args.output}")
    print(f"Holdout: last {HOLDOUT_WEEKS} weeks ({weeks[-HOLDOUT_WEEKS].isoformat()} to {weeks[-1].isoformat()})")
    for row in summary:
        print(f"  {row['forecastability']:<9} {row['sku_count']:>3} SKU(s)  avg MAPE {row['avg_best_mape']}%")
    wins = {m: sum(1 for r in league if r["best_method"] == m) for m in METHODS}
    print("Method wins: " + ", ".join(f"{m} {n}" for m, n in wins.items()))


if __name__ == "__main__":
    main()
