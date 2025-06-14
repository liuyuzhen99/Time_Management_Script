import pandas as pd
from BiDirectionalMap import BiDirectionalMap


class SearchClass:
    def __init__(self):
        self.path = ''
        self.df = pd.DataFrame()
        self.bi_map = BiDirectionalMap()
        self.feedback = True

    def find_search_result(self, staff_id):
        if staff_id:
            log_list = []
            df = pd.read_excel(self.path, sheet_name='Sheet 2')
            for i in range(len(df)):
                if int(df.iloc[i, 0]) == int(staff_id):
                    log_list.append({'Staff ID': df.iloc[i, 0],
                                     'Name': df.iloc[i, 1],
                                     'Date': df.iloc[i, 2],
                                     'Clock in': df.iloc[i, 3],
                                     'Clock out': df.iloc[i, 4],
                                     'Working Hour': df.iloc[i, 5],
                                     'Status': df.iloc[i, 6],
                                     'Department': df.iloc[i, 7]})
            if len(log_list) == 0:
                self.feedback = False
            else:
                self.df = pd.DataFrame(log_list)
        else:
            self.feedback = False
