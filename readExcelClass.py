import pandas as pd


class ReadExcelClass:
    def __init__(self, path=''):
        self.path = path
        self.dataframe = ''

    def read__file(self):
        self.dataframe = pd.read_excel(self.path)
