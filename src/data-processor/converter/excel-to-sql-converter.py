import sqlite3
import pandas as pd
import os

class ExcelToSQLConverter():
    def __init__(self, xl_filename):
        self.xl_filename = xl_filename
        self.xl_filepath = os.path.join("data/excel", "clean_" + xl_filename)
        self.db_folder = os.path.normpath("data/db")
        self.db_filename = "data.db"
        self.db_filepath = os.path.join(self.db_folder, self.db_filename)
        self.connection = sqlite3.connect(self.db_filepath)
        self.cursor = self.connection.cursor

    def _read_from_excel(self) -> pd.DataFrame:
        dataframe = pd.read_excel(self.xl_filepath, sheet_name = "Sheet1")
        return dataframe

    @staticmethod
    def _clean_numerical_columns(dataframe):
        col_names = list(dataframe.columns)
        non_num_cols = ["Postcode District", "Bedroom Category"]
        num_cols = [col for col in col_names if col not in non_num_cols]
        dataframe[num_cols] = dataframe[num_cols].apply(pd.to_numeric,
                                                        axis=0, 
                                                        raw=False, 
                                                        args=('coerce', 'integer'), 
                                                        result_type='expand').astype("Int64")
        return dataframe
    
    def convert(self):
        df = self._read_from_excel()
        df = self._clean_numerical_columns(df)
        df.to_sql(self.db_filepath, self.connection, if_exists="replace")

if __name__ == "__main__":
    ExcelToSQLConverter("londonrentalstatsaccessibleq32024.xlsx").convert()
