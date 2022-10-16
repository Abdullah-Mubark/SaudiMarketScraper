import os
import logging
import json
from datetime import datetime, timedelta
from functools import lru_cache

from dacite import from_dict, Config as DaciteConfig

from config import config, Stock, StockData

class StorageManager:

    def __init__(self) -> None:
        self.__folder_name = config.storage.folder_name
        self.__create_folder_if_not_exists()

        self.__existing_files, self.__latest_file = self.__get_existing_files()
        self.__files_to_keep = list(map(lambda date : date.strftime('%Y-%m-%d') + '.json', self.__get_last_n_days(config.storage.keep_last_files_count))) 
        self.__remove_old_files()

    
    def store_stock_date(self, stock_data: StockData) -> StockData:
        file_name = stock_data.date.strftime('%Y-%m-%d') + '.json'
        with open(os.path.join(self.__folder_name, file_name), 'w') as f:
            json = stock_data.to_json()
            f.write(json)
            logging.info(f'Stock data stored in file {file_name}')
        return self.__get_stock_date_from_file(file_name)

    @lru_cache(maxsize=1)
    def get_stock_data_by_date(self, date: datetime) -> StockData:
        file_name = date.strftime('%Y-%m-%d') + '.json'
        return self.__get_stock_date_from_file(file_name)

    def __get_stock_date_from_file(self, file_name) -> StockData:
        try:
            file = open(os.path.join(self.__folder_name, file_name), 'r')
            data = json.load(file)
            return from_dict(data_class=StockData, data=data, config=DaciteConfig(check_types=False))
        except Exception as e:
            logging.error(f'Error while fetching stock file date [{file_name}]. error : {e}')
            return None

    def __get_last_n_days(self, n) -> list[datetime]:
        return [datetime.today() - timedelta(days=i) for i in range(n)]

    def __create_folder_if_not_exists(self) -> None:
        if not os.path.exists(self.__folder_name):
            os.makedirs(self.__folder_name)
            logging.info(f'Created folder {self.__folder_name}')

    def __get_existing_files(self) -> tuple[list[str], str]:
        files = list(filter(lambda file: file.endswith('.json') and len(file.split('-')) == 3, os.listdir(self.__folder_name)))
        files.sort(reverse=True)
        return files, files[0] if len(files) > 0 else None

    def __remove_old_files(self) -> None:
        for file in self.__existing_files:
            if file not in self.__files_to_keep and file != self.__latest_file:
                os.remove(os.path.join(self.__folder_name, file))
                logging.info(f'Removed file {file}')
