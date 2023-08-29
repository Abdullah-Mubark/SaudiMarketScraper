import dataclasses
import json
import os
from dataclasses import dataclass
from datetime import date, datetime

import yaml
from dacite import from_dict


@dataclass
class Tadawul:
    url: str
    financial_indicators_url: str
    session_key: str
    stocks: list[int]


@dataclass
class Finbox:
    url: str
    fair_value_query: str
    asset_benchmark_query: str
    cookies: dict


@dataclass
class Storage:
    keep_last_files_count: int
    folder_name: str


@dataclass
class Portfolio:
    folder_name: str
    file_name: str


@dataclass
class Excel:
    folder_name: str
    stocks_analysis_file_name: str
    market_analysis_file_name: str
    stocks_analysis_sheet_name: str
    market_analysis_sheet_name: str
    stocks_table_name: str
    stocks_table_start_row: int
    market_table_start_row: int


@dataclass
class Config:
    @classmethod
    def from_yaml(cls, file_path: str):
        with open(file_path, 'r') as f:
            config_dict = yaml.safe_load(f)

        return from_dict(data_class=cls, data=config_dict)

    tadawul: Tadawul
    finbox: Finbox
    storage: Storage
    portfolio: Portfolio
    excel: Excel


@dataclass
class Dividend:
    announcement_date: datetime
    due_date: datetime
    distribution_date: datetime
    distribution_way: str
    amount: float


@dataclass
class Benchmark:
    div_yield: float = None
    p_e: float = None
    p_b: float = None


@dataclass
class FairValue:
    average: float
    uncertainty: str


@dataclass
class Stock:
    name: str = None
    sector: str = None
    industry_group: str = None
    code: int = None
    price: float = None
    _52_week_high: float = None
    _52_week_low: float = None
    market_cap: float = None
    market_cap_percentage: float = None
    issued_shares: float = None
    net_profit: float = None
    shareholders_equity: float = None
    earnings_per_share: float = None
    book_value_per_share: float = None
    quantity_owned: int = None
    cost_price: float = None
    dividends: list[Dividend] = None
    benchmark: Benchmark = None
    fair_value: FairValue = None
    scraped_at: datetime = None
    success_scraping: bool = None


@dataclass
class StockData:
    all_success: bool
    date: date
    stocks: list[Stock]

    def to_json(self):
        return json.dumps(dataclasses.asdict(self), indent=4, default=str)


@dataclass
class FinancialIndicators:
    date: date
    stocks: list[Stock]
    industry_groups: list[str]
    headers: list[str]


config = Config.from_yaml(file_path=os.path.join(os.path.dirname(__file__), 'config.yaml'))
