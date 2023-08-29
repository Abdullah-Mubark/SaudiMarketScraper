import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

logging.basicConfig(
    handlers=[
        logging.StreamHandler(),
        RotatingFileHandler('app.log', maxBytes=1024 * 1024, backupCount=5)
    ],
    level=logging.INFO,
    format='[%(levelname)s] [%(asctime)s] %(message)s',
)

from scraper import StockScraper
from storage_manager import StorageManager
from portfolio_manager import PortfolioManager
from excel_manager import ExcelManager

from config import config, Stock, StockData

stock_scraper = StockScraper()
storage_manager = StorageManager()
portfolio_manager = PortfolioManager()
excel_manager = ExcelManager()


def get_today_stock_data() -> StockData:
    today_stock_data = storage_manager.get_stock_data_by_date(datetime.today())
    stocks_to_scrape = config.tadawul.stocks

    if today_stock_data is None:
        stocks = stock_scraper.scrape_list(stocks_to_scrape)
    elif today_stock_data.all_success:
        stocks = [stock for stock in today_stock_data.stocks if stock.code in stocks_to_scrape]
    else:
        stocks: list[Stock] = []
        for stock_to_scrape in stocks_to_scrape:
            stock_obj = next((s for s in today_stock_data.stocks if s.code == stock_to_scrape), None)
            if stock_obj is not None and stock_obj.success_scraping:
                stocks.append(stock_obj)
            else:
                stock = stock_scraper.scrape_one(stock_to_scrape)
                stocks.append(stock)

    all_success = all(map(lambda stock: stock.success_scraping, stocks))

    return StockData(all_success=all_success, date=datetime.today().date(), stocks=stocks)


def update_stock_data_from_portfolio(stock_data: StockData) -> StockData:
    if portfolio_manager.portfolio_data is None:
        logging.warning('Portfolio data is not loaded')
        return stock_data

    for stock in stock_data.stocks:
        if stock.code in portfolio_manager.portfolio_data:
            quantity_owned, cost_price = portfolio_manager.portfolio_data[stock.code]
            stock.quantity_owned = quantity_owned
            stock.cost_price = cost_price
        else:
            logging.warning(f'Portfolio data does not contain stock with code {stock.code}')

    return stock_data


if __name__ == "__main__":
    stock_data = get_today_stock_data()
    stock_data = update_stock_data_from_portfolio(stock_data)
    stock_data = storage_manager.store_stock_date(stock_data)

    excel_manager.create_stocks_analysis(stock_data)
    excel_manager.create_market_analysis(financial_indicators=stock_scraper.scrape_financial_indicators())
