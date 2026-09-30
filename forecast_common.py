"""Shared constants and data models for the demand forecasting demo.

Used by generate_data.py, validate_data.py, and forecast_demand.py so all
three scripts agree on what a valid SKU record and a valid weekly sales
record look like, and on where the files live.
"""

import math
from datetime import date

from pydantic import BaseModel, Field

CATEGORIES = [
    "Beverages",
    "Snacks",
    "Household",
    "Personal Care",
    "Garden",
    "Toys",
    "Stationery",
    "Pet Supplies",
]

# Common target service levels used across the demo dataset.
SERVICE_LEVELS = [0.90, 0.95, 0.975, 0.99]

SKU_MASTER_RAW_PATH = "data/sku_master_raw.csv"
SKU_MASTER_VALIDATED_PATH = "data/sku_master_validated.csv"
SALES_RAW_PATH = "data/weekly_sales_raw.csv"
SALES_VALIDATED_PATH = "data/weekly_sales_validated.csv"
VALIDATION_REPORT_PATH = "data/validation_report.txt"
OUTPUT_WORKBOOK_PATH = "output/demand_forecast.xlsx"

SKU_FIELDS = [
    "sku",
    "item_name",
    "category",
    "unit_price",
    "lead_time_weeks",
    "service_level",
]

SALES_FIELDS = [
    "sku",
    "week_start",
    "units_sold",
    "on_promo",
]

# Forecasting parameters shared by the generator and the forecaster.
SEASON_LENGTH = 52  # weeks per seasonal cycle
HOLDOUT_WEEKS = 13  # the last quarter of history is held back to score the methods
FORECAST_HORIZON_WEEKS = 4  # "next month"
MOVING_AVERAGE_WINDOW = 8

# Forecastability bands, applied to the best method's holdout MAPE.
# (label, upper bound on MAPE as a fraction)
FORECASTABILITY_BANDS = [
    ("Easy", 0.15),
    ("Moderate", 0.30),
    ("Hard", math.inf),
]


class SkuRecord(BaseModel):
    """Static attributes of one SKU: what it is and how it is replenished."""

    sku: str = Field(min_length=1)
    item_name: str = Field(min_length=1)
    category: str = Field(min_length=1)
    unit_price: float = Field(gt=0)
    lead_time_weeks: int = Field(gt=0)
    service_level: float = Field(gt=0, lt=1)

    class Config:
        extra = "ignore"


class SalesRecord(BaseModel):
    """One SKU's sales for one week (weeks start on a Monday)."""

    sku: str = Field(min_length=1)
    week_start: date
    units_sold: int = Field(ge=0)
    on_promo: bool

    class Config:
        extra = "ignore"
