import pandas as pd
from datetime import datetime
import os
from exception_handler import exception_signal_instance
from MatchException import MatchError
import re
from readExcelClass import ReadExcelClass
from BiDirectionalMap import BiDirectionalMap
from datetime import datetime, timedelta


class BindClass:
    def __init__(self, month, log_record_reader=ReadExcelClass(), staff_li_reader=ReadExcelClass(),
                 leave_li_reader=ReadExcelClass(), bu_trip_reader=ReadExcelClass(), ooo_reader=ReadExcelClass(),
                 card_problem_reader=ReadExcelClass()):
        self.log_record = log_record_reader
        self.staff_li = staff_li_reader
        self.leave_li = leave_li_reader
        self.bu_trip_li = bu_trip_reader
        self.ooo_li = ooo_reader
        self.card_li = card_problem_reader
        self.month = month

    def add_row2df(self, df, target_df, j):
        new_row = pd.DataFrame([df.iloc[j, 0], df.iloc[j, 1],
                                df.iloc[j, 6], df.iloc[j, 2],
                                df.iloc[j, 3], df.iloc[j, 4],
                                df.iloc[j, 5], df.iloc[j, 7]],
                               index=['Staff no.', 'Name with abnormal records', 'Date',
                                      'First Clock in', 'last Clock out', 'working hours',
                                      'abnormal case', 'Department'])
        return pd.concat([target_df, new_row], ignore_index=True)

    def addrow2sumDf(self, df, target_df, j):
        new_row = pd.DataFrame([df.iloc[j, 0], df.iloc[j, 1],
                                df.iloc[j, 6], df.iloc[j, 2],
                                df.iloc[j, 3], df.iloc[j, 4],
                                df.iloc[j, 5], df.iloc[j, 7]],
                               index=['Staff id', 'Name', 'Date', 'Clock in', 'Clock out', 'working Hour', 'Status',
                                      'Department'])
        return pd.concat([target_df, new_row], ignore_index=True)

    def extract_date(self, date):
        # 从log_record的Date/Time提取date 2024-04-17 18:08:20
        # datetime_obj = datetime.strptime(date, "%Y-%m-%d %H:%M:%S")
        datetime_obj = datetime.strptime(date, "%d.%m.%Y %H:%M:%S")
        return int(datetime_obj.day)

    def extract_date_(self, date):
        # 从log_record的Date/Time提取date 2024-04-17 18:08:20
        # datetime_obj = datetime.strptime(date, "%Y-%m-%d")
        datetime_obj = datetime.strptime(date, "%d.%m.%Y")
        return int(datetime_obj.day)

    def extract_time(self, time):
        # 从log_record的Date/Time提取time
        #datetime_obj = datetime.strptime(time, "%Y-%m-%d %H:%M:%S")
        datetime_obj = datetime.strptime(time, "%d.%m.%Y %H:%M:%S")
        return datetime_obj.time()

    def extract_time_(self, time):
        # 从log_record的Date/Time提取time
        #datetime_obj = datetime.strptime(time, "%H:%M:%S")
        datetime_obj = datetime.strptime(time, "%H:%M:%S")
        return datetime_obj.time()

    def workingHour_cal(self, start_time, end_time):
        datetime_obj1 = datetime.strptime(start_time, "%H:%M:%S")
        datetime_obj2 = datetime.strptime(end_time, "%H:%M:%S")
        time_diff = datetime_obj2 - datetime_obj1
        return round(time_diff.total_seconds() / 3600, 2)

    def convert_excel(self):
        columns = ['Staff id', 'Name', 'Clock in', 'Clock out', 'working Hour', 'Status', 'Date', 'Department']
        dataframe_list = [pd.DataFrame(columns=columns) for _ in range(31)]
        bi_map = BiDirectionalMap()
        # 初始化每一个sheet的第一列staff id和第二列Name
        for i in range(31):
            dataframe_list[i]['Staff id'] = self.staff_li.dataframe['Staff Code']
            dataframe_list[i]['Name'] = self.staff_li.dataframe['Last Name'].astype(str) + ' ' + self.staff_li.dataframe['First Name'].astype(str)
            dataframe_list[i]['Date'] = i + 1
            dataframe_list[i]['Department'] = self.staff_li.dataframe['Home Org Unit Short Unit']
        # 创建字典，用来使用staff id快速对应到索引（复杂度为O(1))
        for i in range(len(self.staff_li.dataframe)):
            bi_map.add_mapping(i, str(self.staff_li.dataframe.iloc[i, 0]))

        # 读取leave/Business trip/out of office/card problem文件，将对应的状态更新到对应员工对应天的数据中
        # leave list
        for i in range(len(self.leave_li.dataframe)):
            start_date = datetime.strptime(str(self.leave_li.dataframe.iloc[i, 15]), "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(str(self.leave_li.dataframe.iloc[i, 16]), "%Y-%m-%d %H:%M:%S")
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                if current_date.month == self.month:
                    date_list.append(current_date.day)
                current_date += timedelta(days=1)
            if len(date_list) != 0:
                start = date_list[0]
                end = date_list[len(date_list) - 1]
                staff_id = self.leave_li.dataframe.iloc[i, 0]
                index = bi_map.get_index_by_staff(str(staff_id))
                if index == 0 or index:
                    for j in range(start - 1, end):
                        dataframe_list[j].iloc[index, 5] = 'on leave'

        # Business trip
        for i in range(len(self.bu_trip_li.dataframe)):
            # start = self.extract_date_(str(self.bu_trip_li.dataframe.iloc[i, 2]))
            # end = self.extract_date_(str(self.bu_trip_li.dataframe.iloc[i, 3]))
            # staff_id = self.bu_trip_li.dataframe.iloc[i, 1]
            # index = bi_map.get_index_by_staff(str(staff_id))
            # if index == 0 or index:
            #     for j in range(start - 1, end):
            #         dataframe_list[j].iloc[index, 5] = 'on business trip'
            start_date = datetime.strptime(str(self.bu_trip_li.dataframe.iloc[i, 2]), "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(str(self.bu_trip_li.dataframe.iloc[i, 3]), "%Y-%m-%d %H:%M:%S")
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                if current_date.month == self.month:
                    date_list.append(current_date.day)
                current_date += timedelta(days=1)
            if len(date_list) != 0:
                start = date_list[0]
                end = date_list[len(date_list) - 1]
                staff_id = self.bu_trip_li.dataframe.iloc[i, 1]
                index = bi_map.get_index_by_staff(str(staff_id))
                if index == 0 or index:
                    for j in range(start - 1, end):
                        dataframe_list[j].iloc[index, 5] = 'on business trip'

        # out of office
        for i in range(len(self.ooo_li.dataframe)):
            # start = self.extract_date_(str(self.ooo_li.dataframe.iloc[i, 2]))
            # end = self.extract_date_(str(self.ooo_li.dataframe.iloc[i, 3]))
            # staff_id = self.ooo_li.dataframe.iloc[i, 0]
            # index = bi_map.get_index_by_staff(str(staff_id))
            # if index == 0 or index:
            #     for j in range(start - 1, end):
            #         dataframe_list[j].iloc[index, 5] = 'out of office'
            start_date = datetime.strptime(str(self.ooo_li.dataframe.iloc[i, 2]), "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(str(self.ooo_li.dataframe.iloc[i, 3]), "%Y-%m-%d %H:%M:%S")
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                if current_date.month == self.month:
                    date_list.append(current_date.day)
                current_date += timedelta(days=1)
            if len(date_list) != 0:
                start = date_list[0]
                end = date_list[len(date_list) - 1]
                staff_id = self.ooo_li.dataframe.iloc[i, 0]
                index = bi_map.get_index_by_staff(str(staff_id))
                if index == 0 or index:
                    for j in range(start - 1, end):
                        dataframe_list[j].iloc[index, 5] = 'out of office'

        # card problem
        for i in range(len(self.card_li.dataframe)):
            # start = self.extract_date_(str(self.card_li.dataframe.iloc[i, 2]))
            # end = self.extract_date_(str(self.card_li.dataframe.iloc[i, 3]))
            # staff_id = self.card_li.dataframe.iloc[i, 0]
            # index = bi_map.get_index_by_staff(str(staff_id))
            # if index == 0 or index:
            #     for j in range(start - 1, end):
            #         dataframe_list[j].iloc[index, 5] = 'card problem'
            start_date = datetime.strptime(str(self.card_li.dataframe.iloc[i, 2]), "%Y-%m-%d %H:%M:%S")
            end_date = datetime.strptime(str(self.card_li.dataframe.iloc[i, 3]), "%Y-%m-%d %H:%M:%S")
            date_list = []
            current_date = start_date
            while current_date <= end_date:
                if current_date.month == self.month:
                    date_list.append(current_date.day)
                current_date += timedelta(days=1)
            if len(date_list) != 0:
                start = date_list[0]
                end = date_list[len(date_list) - 1]
                staff_id = self.card_li.dataframe.iloc[i, 0]
                index = bi_map.get_index_by_staff(str(staff_id))
                if index == 0 or index:
                    for j in range(start - 1, end):
                        dataframe_list[j].iloc[index, 5] = 'card problem'

        # 读取log record
        for i in range(len(self.log_record.dataframe)):
            if not pd.isnull(self.log_record.dataframe.iloc[i, 3]):
                index = bi_map.get_index_by_staff(str(round(self.log_record.dataframe.iloc[i, 3]))) #7
            if index == 0 or index:
                date = self.extract_date(str(self.log_record.dataframe.iloc[i, 0])) - 1
                time = self.extract_time(str(self.log_record.dataframe.iloc[i, 0]))
                if time < datetime.strptime("13:00:00", "%H:%M:%S").time():
                    if pd.isnull(dataframe_list[date].iloc[index, 2]) or time < datetime.strptime(
                            dataframe_list[date].iloc[index, 2], "%H:%M:%S").time():  # to be change
                        dataframe_list[date].iloc[index, 2] = str(time)
                else:
                    if pd.isnull(dataframe_list[date].iloc[index, 3]) or time > datetime.strptime(
                            dataframe_list[date].iloc[index, 3], "%H:%M:%S").time():
                        dataframe_list[date].iloc[index, 3] = str(time)

        # 计算每行数据的working hour和Status, 并生成abnormal sheet
        abnormal_data = []  # pd.DataFrame(columns=['Staff no.', 'Name with abnormal records', 'Date', 'First Clock in', 'last Clock out', 'working hours', 'abnormal case', 'Department'])
        sum_data = []  # pd.DataFrame(columns=['Staff id', 'Name', 'Date',  'Clock in', 'Clock out', 'working Hour', 'Status', 'Department'])
        for j in range(len(dataframe_list[0])):
            for i in range(31):
                if pd.isnull(dataframe_list[i].iloc[j, 5]):
                    if not pd.isnull(dataframe_list[i].iloc[j, 2]) and not pd.isnull(dataframe_list[i].iloc[j, 3]):
                        dataframe_list[i].iloc[j, 4] = self.workingHour_cal(
                            dataframe_list[i].iloc[j, 2] , dataframe_list[i].iloc[j, 3])
                        # print(dataframe_list[i].iloc[j, 2])
                        if self.extract_time_(dataframe_list[i].iloc[j, 2]) > datetime.strptime("9:00:00",
                                                                                               "%H:%M:%S").time():
                            dataframe_list[i].iloc[j, 5] = "come late"
                            abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                                  'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                                  'Date': dataframe_list[i].iloc[j, 6],
                                                  'First Clock in': dataframe_list[i].iloc[j, 2],
                                                  'last Clock out': dataframe_list[i].iloc[j, 3],
                                                  'working hours': dataframe_list[i].iloc[j, 4],
                                                  'abnormal case': dataframe_list[i].iloc[j, 5],
                                                  'Department': dataframe_list[i].iloc[j, 7]})
                        elif self.extract_time_(dataframe_list[i].iloc[j, 3]) < datetime.strptime("17:00:00",
                                                                                                 "%H:%M:%S").time():
                            dataframe_list[i].iloc[j, 5] = "leave early"
                            abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                                  'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                                  'Date': dataframe_list[i].iloc[j, 6],
                                                  'First Clock in': dataframe_list[i].iloc[j, 2],
                                                  'last Clock out': dataframe_list[i].iloc[j, 3],
                                                  'working hours': dataframe_list[i].iloc[j, 4],
                                                  'abnormal case': dataframe_list[i].iloc[j, 5],
                                                  'Department': dataframe_list[i].iloc[j, 7]})
                        else:
                            dataframe_list[i].iloc[j, 5] = "normal"
                    elif pd.isnull(dataframe_list[i].iloc[j, 2]) and not pd.isnull(dataframe_list[i].iloc[j, 3]):
                        dataframe_list[i].iloc[j, 5] = "clock in missing"
                        abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                              'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                              'Date': dataframe_list[i].iloc[j, 6],
                                              'First Clock in': dataframe_list[i].iloc[j, 2],
                                              'last Clock out': dataframe_list[i].iloc[j, 3],
                                              'working hours': dataframe_list[i].iloc[j, 4],
                                              'abnormal case': dataframe_list[i].iloc[j, 5],
                                              'Department': dataframe_list[i].iloc[j, 7]})
                    elif not pd.isnull(dataframe_list[i].iloc[j, 2]) and pd.isnull(dataframe_list[i].iloc[j, 3]):
                        dataframe_list[i].iloc[j, 5] = "clock out missing"
                        abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                              'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                              'Date': dataframe_list[i].iloc[j, 6],
                                              'First Clock in': dataframe_list[i].iloc[j, 2],
                                              'last Clock out': dataframe_list[i].iloc[j, 3],
                                              'working hours': dataframe_list[i].iloc[j, 4],
                                              'abnormal case': dataframe_list[i].iloc[j, 5],
                                              'Department': dataframe_list[i].iloc[j, 7]})
                    else:
                        dataframe_list[i].iloc[j, 5] = "no log record"
                        abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                              'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                              'Date': dataframe_list[i].iloc[j, 6],
                                              'First Clock in': dataframe_list[i].iloc[j, 2],
                                              'last Clock out': dataframe_list[i].iloc[j, 3],
                                              'working hours': dataframe_list[i].iloc[j, 4],
                                              'abnormal case': dataframe_list[i].iloc[j, 5],
                                              'Department': dataframe_list[i].iloc[j, 7]})
                else:
                    abnormal_data.append({'Staff no.': dataframe_list[i].iloc[j, 0],
                                          'Name with abnormal records': dataframe_list[i].iloc[j, 1],
                                          'Date': dataframe_list[i].iloc[j, 6],
                                          'First Clock in': dataframe_list[i].iloc[j, 2],
                                          'last Clock out': dataframe_list[i].iloc[j, 3],
                                          'working hours': dataframe_list[i].iloc[j, 4],
                                          'abnormal case': dataframe_list[i].iloc[j, 5],
                                          'Department': dataframe_list[i].iloc[j, 7]})
                sum_data.append({'Staff id': dataframe_list[i].iloc[j, 0],
                                 'Name': dataframe_list[i].iloc[j, 1],
                                 'Date': dataframe_list[i].iloc[j, 6],
                                 'Clock in': dataframe_list[i].iloc[j, 2],
                                 'Clock out': dataframe_list[i].iloc[j, 3],
                                 'working Hour': dataframe_list[i].iloc[j, 4],
                                 'Status': dataframe_list[i].iloc[j, 5],
                                 'Department': dataframe_list[i].iloc[j, 7]})

        # 生成最后的excel表格
        with pd.ExcelWriter(os.path.dirname(self.log_record.path) + '\\Access_logbook.xlsx', engine='xlsxwriter') as writer:
            pd.DataFrame(abnormal_data).to_excel(writer, sheet_name='Sheet 1', index=False)
            pd.DataFrame(sum_data).to_excel(writer, sheet_name='Sheet 2', index=False)

