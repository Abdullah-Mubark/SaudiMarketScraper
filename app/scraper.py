import bs4
from bs4 import BeautifulSoup
import requests

from datetime import datetime
import logging
import json
from enum import Enum
from functools import lru_cache

from config import config, Stock, Dividend, Benchmark, FairValue, FinancialIndicators


class FinancialIndicatorsColumnsIndex(Enum):
    INDUSTRY_GROUP = 0
    COMPANY = 0
    PRICE = 1
    ISSUED_SHARES = 2
    NET_INCOME = 3
    SHAREHOLDERS_EQUITY = 4
    MARKET_CAP = 5
    MARKET_CAP_PERCENTAGE = 6
    EARNINGS_PER_SHARE = 7
    P_E_RATIO = 8
    BOOK_VALUE = 9
    P_B_RATIO = 10



class StockScraper:
    def __init__(self) -> None:
        self.tadawul_stock_scraping_url = f'{config.tadawul.url}/{config.tadawul.session_key}'
        self.financial_indicators_url = config.tadawul.financial_indicators_url
        self.finbox_fair_value_scraping_url = f'{config.finbox.url}/v5/query?raw=true'
        self.finbox_fair_value_query = config.finbox.fair_value_query
        self.finbox_asset_benchmark_query = config.finbox.asset_benchmark_query
        self.finbox_cookies = config.finbox.cookies
        self.unkown_stock_name = 'Unknown'

    def scrape_list(self, stocks: list[int]) -> list[Stock]:
        return [self.__scrape_stock_info(self.tadawul_stock_scraping_url, stock_code) for stock_code in stocks]

    def scrape_one(self, stock_code: int) -> Stock:
        return self.__scrape_stock_info(self.tadawul_stock_scraping_url, stock_code)

    @lru_cache(maxsize=1)
    def scrape_financial_indicators(self) -> FinancialIndicators:
        scrapping_url = self.financial_indicators_url
        response = requests.get(scrapping_url)

        if response.status_code != 200:
            logging.error(f'Failed to scrape financial indicators')
            return None
        
        data = BeautifulSoup(response.text, 'html.parser')
        logging.info(f'Financial indicators scraped successfully')

        table = data.find('table', {'class': 'Table3'})
        if table == None:
            logging.error(f'Failed to find financial indicators table')
            return None

        try:
            rows = table.find_all('tr')
            table_headers = [header.text.strip() for header in rows[0].find_all('th')]
            stocks: list[Stock] = []
            industry_groups: list[str] = []

            for row in rows[1:]:
                row_data = {table_headers[i]: cell.text.strip() for i, cell in enumerate(row.find_all('td'))}
                is_industry_group_row = row_data[table_headers[FinancialIndicatorsColumnsIndex.PRICE.value]] == ''
    
                if is_industry_group_row:
                    industry_group_name = row_data[table_headers[FinancialIndicatorsColumnsIndex.INDUSTRY_GROUP.value]]
                    industry_groups.append(industry_group_name)
                    for i in range(len(stocks) - 1, -1, -1):
                        if stocks[i].industry_group == None:
                            stocks[i].industry_group = industry_group_name
                        else:
                            break
                else:
                    stock = Stock(
                        name=row_data[table_headers[FinancialIndicatorsColumnsIndex.COMPANY.value]],
                        price=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.PRICE.value]]),
                        issued_shares=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.ISSUED_SHARES.value]].replace(',', '')),
                        net_profit=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.NET_INCOME.value]].replace(',', '')),
                        shareholders_equity=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.SHAREHOLDERS_EQUITY.value]].replace(',', '')),
                        market_cap=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.MARKET_CAP.value]].replace(',', '')),
                        market_cap_percentage=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.MARKET_CAP_PERCENTAGE.value]]),
                        earnings_per_share=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.EARNINGS_PER_SHARE.value]]),
                        book_value_per_share=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.BOOK_VALUE.value]]),
                        benchmark=Benchmark(
                                p_e=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.P_E_RATIO.value]]),
                                p_b=self.__cast_to_float(row_data[table_headers[FinancialIndicatorsColumnsIndex.P_B_RATIO.value]])
                                )
                    )
                    stocks.append(stock)
        except Exception as e:
            logging.exception(f'Failed to parse financial indicators table. Error: {e}')
            return None
        
        logging.info(f'Financial indicators parsed successfully')
        logging.info(f'Financial indicators Industry groups: {industry_groups}')

        table_headers.insert(0, 'Industry Group')

        return FinancialIndicators(
            industry_groups=industry_groups,
            stocks=stocks,
            headers=table_headers,
            date=datetime.today().date()
        )
                    

    def __scrape_stock_info(self, tadawul_scraping_url: str, stock_code: int) -> Stock:
        stock = Stock(code=stock_code, scraped_at=datetime.now())

        scraping_url = f'{tadawul_scraping_url}/?companySymbol={stock_code}'
        response = requests.get(scraping_url)

        if response.status_code != 200:
            logging.error(f'Failed to scrape stock info [{stock_code}]')
            stock.success_scraping = False
            return stock
        
        data = BeautifulSoup(response.text, 'html.parser')
        logging.info(f'Stock info scraped successfully [{stock_code}]')

        stock.name = self.__extract_stock_name(data, stock_code)
        logging.info(f'Stock [{stock_code}] name is {stock.name}')

        stock.sector = self.__extract_stock_sector(data, stock_code)
        logging.info(f'Stock [{stock.name}] sector is {stock.sector}')

        stock.industry_group = self.__extract_stock_industry_group(data, stock_code)
        logging.info(f'Stock [{stock.name}] industry group is {stock.industry_group}')
        
        stock.price = self.__extract_stock_price(data, stock_code)
        logging.info(f'Stock price for {stock.name} is {stock.price}')

        stock._52_week_high, stock._52_week_low = self.__extract_52_week_prices(data, stock_code)
        logging.info(f'52 week max/min prices for {stock.name} is {stock._52_week_high}/{stock._52_week_low}')

        stock.dividends = self.__extract_stock_dividends(data, stock_code)
        if stock.dividends:
            logging.info(f'Extracted {len(stock.dividends)} dividends for {stock.name}')

        stock.benchmark = self.__scrape_stock_benchmark(stock_code)

        stock.fair_value = self.__scrape_fair_value(stock_code)
        if stock.fair_value:
            logging.info(f'Fair value for {stock.name} is {stock.fair_value.average}')

        self.__update_stocks_info_from_financial_indicators(stock)

        if stock.name == self.unkown_stock_name or stock.sector == None or stock.industry_group == None or stock.price < 0 or stock._52_week_high < 0 or stock._52_week_low < 0 \
             or stock.dividends == None or stock.fair_value == None:
            stock.success_scraping = False
        else:
            stock.success_scraping = True
        
        return stock


    def __extract_stock_name(self, data: bs4.BeautifulSoup, stock_code: int) -> str:
        try:
            return data.find('p', {'class': 'trade_name'}).find('strong').text.strip()
        except Exception as e:
            logging.exception(f'Failed to extract stock name for stock [{stock_code}]. Error: {e}')
            return self.unkown_stock_name

    def __extract_stock_sector(self, data: bs4.BeautifulSoup, stock_code: int) -> str:
        try:
            return data.find_all('p', {'class': 'trade_name'})[1].find('strong').text.strip()
        except Exception as e:
            logging.exception(f'Failed to extract stock sector for stock [{stock_code}]. Error: {e}')
            return None

    def __extract_stock_industry_group(self, data: bs4.BeautifulSoup, stock_code: int) -> str:
        try:
            return data.find_all('p', {'class': 'trade_name'})[1].find('a').text.strip()
        except Exception as e:
            logging.exception(f'Failed to extract industry group for stock [{stock_code}]. Error: {e}')
            return None

    def __extract_stock_price(self, data: bs4.BeautifulSoup, stock_code: int) -> float:
        try:
            return float(data.find('div', {'id': 'chart_tab1'})
                            .find('div' , {'class': 'table_sep'})
                            .find('div')
                            .find('dd')
                            .text.strip())
        except Exception as e:
            logging.exception(f'Failed to extract stock price for stock [{stock_code}]. Error: {e}')
            return -1

    def __extract_52_week_prices(self, data: bs4.BeautifulSoup, stock_code: int) -> tuple[float, float]:
        try:
            stock_52_week = data.find('p', text='52 WEEK').find_next_sibling('table').find('tbody').find('tr').find_all('td')
            stock_52_week_min = float(stock_52_week[1].next.text.strip())
            stock_52_week_max = float(stock_52_week[0].next.text.strip())

            return stock_52_week_max, stock_52_week_min
        except Exception as e:
            logging.exception(f'Failed to extract 52 week prices for stock [{stock_code}]. Error: {e}')
            return -1, -1


    def __extract_stock_dividends(self, data: bs4.BeautifulSoup, stock_code: int) -> list[Dividend]:
        try:
            dividends_table = data.find('table', {'id': 'dividendsTable'})
            dividends_tbody = dividends_table.find('tbody')
            dividends = [self.__extract_dividends_row(dividend_row, stock_code) for dividend_row in dividends_tbody.find_all('tr')]
            return list(filter(lambda dividend: dividend != None, dividends))
        except Exception as e:
            logging.exception(f'Failed to extract stock dividends. Error: {e}')
            return None

    def __extract_dividends_row(self, row: bs4.element.Tag, stock_code: int) -> Dividend:
        date_format = '%Y/%m/%d'
        tds = [td.text.strip() for td in row.find_all('td')] 
        try: 
            return Dividend(
            anouncement_date=datetime.strptime(tds[0], date_format),
            due_date=datetime.strptime(tds[1], date_format),
            distribution_date=datetime.strptime(tds[2], date_format),
            distribution_way=tds[3],
            amount=float(tds[4])
        )
        except Exception as e:
            logging.exception(f'Failed to extract dividend row for stock [{stock_code}]. Error: {e}')
            return None
        
    
    def __scrape_stock_benchmark(self, stock_code: int) -> Benchmark:
        headers =  { 'Content-Type': 'application/json' }
        data = {'query': self.finbox_asset_benchmark_query, 'variables': {'ticker':f'SASE:{stock_code}'}}

        response = requests.post(self.finbox_fair_value_scraping_url, headers=headers, data=json.dumps(data), cookies=self.finbox_cookies)

        if response.status_code != 200:
            logging.error(f'Failed to scrape stock benchmark [{stock_code}]')
            return Benchmark()

        try:
           stats = response.json()['data']['asset']['stats']
           return Benchmark(
                div_yield=float(stats['quote']['div_yield']['company'])
            )
        except Exception as e:
            logging.exception(f'Failed to extract div yield for stock [{stock_code}]. Error: {e}')
            return Benchmark()


    def __scrape_fair_value(self, stock_code: int) -> FairValue:
        headers =  { 'Content-Type': 'application/json' }
        data = {'query': self.finbox_fair_value_query, 'variables': {'ticker':f'SASE:{stock_code}'}}

        response = requests.post(self.finbox_fair_value_scraping_url, headers=headers, data=json.dumps(data))
        if response.status_code != 200:
            logging.error(f'Failed to scrape fair value for [{stock_code}]')
            return None
        
        try:
            fair_value = response.json()['data']['asset']['fair_value']
            return FairValue(average=float(fair_value['averages']['price']), uncertinty=fair_value['uncertainty'])
        except Exception as e:
            logging.exception(f'Failed to extract fair value for stock [{stock_code}]. Error: {e}')
            return None

    def __update_stocks_info_from_financial_indicators(self, stock: Stock):
        financial_indicators: FinancialIndicators = self.scrape_financial_indicators()
        if financial_indicators == None:
            return
        
        stock_from_financial_indicators = next((s for s in financial_indicators.stocks if s.name.lower() == stock.name.lower() and \
                                                                                          s.industry_group.lower() == stock.industry_group.lower() ), None)
        if stock_from_financial_indicators == None:
            logging.error(f'Failed to find stock [{stock.name}] in financial indicators')
            return
        
        stock.issued_shares = stock_from_financial_indicators.issued_shares
        stock.net_profit = stock_from_financial_indicators.net_profit
        stock.shareholders_equity = stock_from_financial_indicators.shareholders_equity
        stock.market_cap = stock_from_financial_indicators.market_cap
        stock.market_cap_percentage = stock_from_financial_indicators.market_cap_percentage
        stock.earnings_per_share = stock_from_financial_indicators.earnings_per_share
        stock.book_value_per_share = stock_from_financial_indicators.book_value_per_share
        stock.benchmark.p_e = stock_from_financial_indicators.benchmark.p_e
        stock.benchmark.p_b = stock_from_financial_indicators.benchmark.p_b
        logging.info(f'Stock [{stock.name}] info was updated from financial indicators')


    def __cast_to_float(self, value: str) -> float:
        try:
            return float(value)
        except ValueError:
            return None
