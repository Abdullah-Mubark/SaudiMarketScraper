import logging
import os

from config import config
import re

class PortfolioManager:

    def __init__(self) -> None:
        self.__folder_name = config.portfolio.folder_name
        self.__file_name = config.portfolio.file_name
        self.portfolio_data = self.__parse_portfolio_file()


    def __fetch_portolio_file(self) -> list[str]:
        try:
            file = open(os.path.join(self.__folder_name, self.__file_name), 'r')
            return file.readlines()
        except Exception as e:
            logging.error(f'Error while fetching portfolio file. error : {e}')
            return None

    
    def __is_number(self, string: str) -> bool:
        try:
            float(string)
            return True
        except ValueError:
            return False
    
    def __parse_portfolio_file(self) -> dict:
        file_content: list[str] = self.__fetch_portolio_file() 
        if file_content is None:
            return None
        
        try:
            portfolio = {}
            numbers_in_quotes_pattern = r'\"(.+?)\"'
            stock_code_col_index = 7
            stock_quantity_col_index = 6
            stock_cost_price_col_index = 4

            for index, line in enumerate(file_content):
                numbers_in_quotes = re.findall(numbers_in_quotes_pattern, line)
                if len(numbers_in_quotes) > 0:
                    numbers_fixed = [number_in_quotes.replace(',', '') for number_in_quotes in numbers_in_quotes]
                    file_content[index] = re.sub(numbers_in_quotes_pattern, '%s', line)  % tuple(numbers_fixed)

            lines_split = [line.split(',') for line in file_content]
            for line in lines_split:
                while '' in line:
                    line.remove('')

            stocks_lines = list(filter(lambda ls: \
                all(map(lambda col : self.__is_number(col), ls))
                    , lines_split))

            for stock_line in stocks_lines:
                stock_code = int(stock_line[stock_code_col_index])
                stock_quantity = int(stock_line[stock_quantity_col_index])
                stock_cost_price = float(stock_line[stock_cost_price_col_index])
                portfolio[stock_code] = (stock_quantity, stock_cost_price)
            
            return portfolio
        except Exception as e:
            logging.error(f'Error while parsing portfolio file. error : {e}')
            return None


