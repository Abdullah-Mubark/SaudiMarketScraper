from datetime import datetime

from config import config, Stock

class Analytics:

    def get_dividends_sum_amount_for_year(self, stock: Stock, year: int) -> float:
        if stock.dividends is None:
            return ''

        dividends = [div for div in stock.dividends if datetime.strptime(div.due_date, "%Y-%m-%d %H:%M:%S").date().year == year]
        return '' if len(dividends) == 0 else sum([div.amount for div in dividends]) 

    def get_portfolio_total_value(self, stocks: list[Stock]) -> float:
        protfolio = [stock for stock in stocks if stock.price is not None and stock.price != -1 and stock.quantity_owned is not None and stock.quantity_owned != -1]
        return sum([stock.price * stock.quantity_owned for stock in protfolio])