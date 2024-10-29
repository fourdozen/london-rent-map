import openpyxl
import os

# Take Standard excel file from Mayor's office data release and keep only relevant data
class ExcelCleaner():
    
    def __init__(self, filename):
        self.filepath = self._data_filepath(filename)

    def _data_filepath(filename):
        cwd = os.getcwd()
        filepath = os.path.join(cwd, "data", filename)
        return filepath