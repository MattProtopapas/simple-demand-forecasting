"""Generate a fake weekly sales history with Faker.

Produces two CSVs that the rest of the pipeline (validate_data.py,
forecast_demand.py) consumes:

  - a SKU master (name, category, price, lead time, service level)
  - three years of weekly unit sales per SKU, built from a base level plus
    trend, yearly seasonality, random promo spikes and noise

A small percentage of sales rows are deliberately corrupted (negative units,
blank fields, bad dates, unknown SKUs, a duplicate week) so validate_data.py
has real work to do - pass --error-rate 0 for a perfectly clean dataset.
"""

import argparse
import csv
import math
import os
import random
from datetime import date, timedelta

from faker import Faker

from forecast_common import (
    CATEGORIES,
    SALES_FIELDS,
    SALES_RAW_PATH,
    SEASON_LENGTH,
    SERVICE_LEVELS,
    SKU_FIELDS,
    SKU_MASTER_RAW_PATH,
)

PRODUCT_NOUNS = ["Pack", "Box", "Bottle", "Bundle", "Set", "Jar", "Bag", "Tin"]


def most_recent_monday(today: date | None = None) -> date:
    today = today or date.today()
    return today - timedelta(days=today.weekday())


def build_item_name(fake: Faker) -> str:
    return f"{fake.word().capitalize()} {fake.word().capitalize()} {random.choice(PRODUCT_NOUNS)}"


def generate_sku(fake: Faker, index: int) -> dict:
    return {
        "sku": f"SKU-{index:03d}",
        "item_name": build_item_name(fake),
        "category": random.choice(CATEGORIES),
        "unit_price": round(random.uniform(1.5, 80), 2),
        "lead_time_weeks": random.randint(1, 6),
        "service_level": random.choice(SERVICE_LEVELS),
    }


def build_demand_profile(intermittent: bool) -> dict:
    """Pick the hidden parameters that drive one SKU's demand curve."""
    base = random.uniform(1.0, 6.0) if intermittent else random.uniform(15, 400)
    return {
        "base": base,
        # weekly trend as a fraction of base: mostly gentle growth, a few decliners
        "trend_per_week": base * random.uniform(-0.0015, 0.004),
        # yearly sine wave amplitude (fraction of base) and phase shift (weeks)
        "season_amplitude": random.uniform(0.05, 0.5),
        "season_phase": random.uniform(0, SEASON_LENGTH),
        # optional extra Q4 bump (fraction of base), peaking around ISO week 48
        "q4_bump": random.choice([0.0, 0.0, 0.4, 0.8, 1.2]),
        # noise as a coefficient of variation
        "noise_cv": random.uniform(0.05, 0.35),
        # promo weeks: probability per week and multiplicative lift
        "promo_rate": random.uniform(0.03, 0.10),
        "promo_lift": random.uniform(1.6, 3.0),
        "intermittent": intermittent,
    }


def seasonal_factor(week_index: int, week_start: date, profile: dict) -> float:
    angle = 2 * math.pi * (week_index + profile["season_phase"]) / SEASON_LENGTH
    factor = 1.0 + profile["season_amplitude"] * math.sin(angle)
    if profile["q4_bump"]:
        iso_week = week_start.isocalendar()[1]
        # bell curve centred on ISO week 48 with a ~3 week spread
        factor += profile["q4_bump"] * math.exp(-((iso_week - 48) ** 2) / (2 * 3.0**2))
    return max(factor, 0.05)


def generate_sales(sku: str, profile: dict, weeks: list[date]) -> list[dict]:
    rows = []
    for i, week_start in enumerate(weeks):
        level = profile["base"] + profile["trend_per_week"] * i
        expected = level * seasonal_factor(i, week_start, profile)

        on_promo = random.random() < profile["promo_rate"]
        if on_promo:
            expected *= profile["promo_lift"]

        if profile["intermittent"]:
            # low-volume items: lumpy, with plenty of zero weeks
            units = sum(1 for _ in range(int(expected * 4)) if random.random() < 0.25)
        else:
            units = random.gauss(expected, expected * profile["noise_cv"])

        rows.append(
            {
                "sku": sku,
                "week_start": week_start.isoformat(),
                "units_sold": max(0, round(units)),
                "on_promo": on_promo,
            }
        )
    return rows


def corrupt_row(row: dict) -> dict:
    """Introduce one realistic data-quality issue into a sales row, in place."""
    issue = random.choice(
        ["negative_units", "blank_units", "bad_date", "bad_promo", "blank_sku", "unknown_sku"]
    )
    if issue == "negative_units":
        row["units_sold"] = -abs(int(row["units_sold"])) - 1
    elif issue == "blank_units":
        row["units_sold"] = ""
    elif issue == "bad_date":
        row["week_start"] = row["week_start"][:8] + "99"  # e.g. 2025-03-99
    elif issue == "bad_promo":
        row["on_promo"] = "maybe"
    elif issue == "blank_sku":
        row["sku"] = ""
    elif issue == "unknown_sku":
        row["sku"] = "SKU-999"
    return row


def generate_dataset(
    sku_count: int, weeks_count: int, end: date, error_rate: float, seed: int | None
) -> tuple[list[dict], list[dict]]:
    if seed is not None:
        random.seed(seed)
        Faker.seed(seed)
    fake = Faker()

    weeks = [end - timedelta(weeks=weeks_count - 1 - i) for i in range(weeks_count)]

    skus = [generate_sku(fake, i) for i in range(1, sku_count + 1)]
    # roughly one SKU in eight is a slow, lumpy mover
    intermittent_count = max(1, sku_count // 8) if sku_count > 1 else 0
    intermittent = set(random.sample(range(sku_count), k=intermittent_count))

    sales: list[dict] = []
    for idx, sku in enumerate(skus):
        profile = build_demand_profile(intermittent=idx in intermittent)
        sales.extend(generate_sales(sku["sku"], profile, weeks))

    n_errors = round(len(sales) * error_rate)
    for row in random.sample(sales, k=min(n_errors, len(sales))):
        corrupt_row(row)

    # Inject a duplicate (sku, week) row to exercise the validator's duplicate
    # check, as long as there's more than one row to work with.
    if error_rate > 0 and len(sales) > 1:
        sales.insert(len(sales) // 2, dict(sales[0]))

    return skus, sales


def write_csv(rows: list[dict], fieldnames: list[str], path: str) -> None:
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate fake weekly sales data for the forecasting pipeline.")
    parser.add_argument("--skus", type=int, default=30, help="Number of SKUs to generate (default: 30)")
    parser.add_argument("--weeks", type=int, default=156, help="Weeks of history to generate (default: 156 = 3 years)")
    parser.add_argument(
        "--end",
        type=date.fromisoformat,
        default=most_recent_monday(),
        help="Monday of the last week of history, ISO format (default: the most recent Monday)",
    )
    parser.add_argument(
        "--error-rate",
        type=float,
        default=0.02,
        help="Fraction of sales rows (0-1) to deliberately corrupt, to exercise validate_data.py (default: 0.02)",
    )
    parser.add_argument("--seed", type=int, default=None, help="Random seed for reproducible output")
    parser.add_argument("--sales-output", type=str, default=SALES_RAW_PATH, help="Path to write the raw sales CSV to")
    parser.add_argument(
        "--master-output", type=str, default=SKU_MASTER_RAW_PATH, help="Path to write the SKU master CSV to"
    )
    args = parser.parse_args()

    end = args.end - timedelta(days=args.end.weekday())  # snap to Monday
    skus, sales = generate_dataset(args.skus, args.weeks, end, args.error_rate, args.seed)

    write_csv(skus, SKU_FIELDS, args.master_output)
    write_csv(sales, SALES_FIELDS, args.sales_output)

    print(f"Generated {len(skus)} SKUs -> {args.master_output}")
    print(
        f"Generated {len(sales)} weekly sales rows ({round(args.error_rate * 100)}% target error rate) "
        f"-> {args.sales_output}"
    )


if __name__ == "__main__":
    main()
