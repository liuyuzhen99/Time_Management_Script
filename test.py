import pandas as pd
from datetime import datetime, timedelta
from bindClass import BindClass
from readExcelClass import ReadExcelClass

log_record_path = r'C:\Users\eqhw2il\Downloads\5月考勤数据 1.xlsx'
leave_path = r'C:\Users\eqhw2il\Downloads\test-leavelist.xlsx'
business_trip_path = r'C:\Users\eqhw2il\Downloads\test-businessTriplist.xlsx'
ooo_path = r'C:\Users\eqhw2il\Downloads\test-ooolist.xlsx'
card_problem_path = r'C:\Users\eqhw2il\Downloads\test-cardProblemlist.xlsx'
employee_path = r'C:\Users\eqhw2il\Downloads\test-Employee list.xlsx'
#
log_df = ReadExcelClass(log_record_path)
leave_df = ReadExcelClass(leave_path)
business_trip_df = ReadExcelClass(business_trip_path)
ooo_df = ReadExcelClass(ooo_path)
card_problem_df = ReadExcelClass(card_problem_path)
employee_df = ReadExcelClass(employee_path)
#
log_df.read__file()
leave_df.read__file()
business_trip_df.read__file()
ooo_df.read__file()
card_problem_df.read__file()
employee_df.read__file()

bindClass = BindClass(5, log_df, employee_df, leave_df, business_trip_df, ooo_df, card_problem_df)
bindClass.convert_excel()
# print(str(employee_df.dataframe.iloc[2, 1]) + str(employee_df.dataframe.iloc[2, 2]))
# print(str(employee_df.dataframe.iloc[2, 2]))
# print(log_df.dataframe.iloc[0, 7])

# path = r'C:\Users\eqhw2il\OneDrive - 奥迪一汽新能源汽车有限公司\桌面\项目相关\TaAMS\Access_logbook.xlsx'
# df = pd.read_excel(path, sheet_name='Sheet 2')
# print(int(df.iloc[1, 0]) == 122772)


