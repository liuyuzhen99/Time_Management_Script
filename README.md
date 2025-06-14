# Time Management Script打卡记录脚本

## Business Background业务背景

Company lack of system for time Management. Need a interim solution to check the clock record.
公司缺少考勤系统，需要一个临时解决方案进行考勤

## Function功能
upload tables to the application上传以下表格到应用:
1. access_logbook.xlsx: staff clock record员工打卡记录
2. businessTriplist.xlsx: staff business trip list员工出差记录
3. cardProblemlist.xlsx: staff card problem list员工卡失效记录
4. leaveList.xlsx: staff leave list员工休假记录
5. ooolist.xlsx: staff out of office list员工外勤记录
6. employeelist.xlsx: staff list员工列表
based on above file to generate clock info for each staff, and support to check the result by using employee id根据以上信息生成每个员工每月的打卡记录，并可根据员工号进行筛选

![image](https://github.com/user-attachments/assets/dbd00063-34e9-4dba-b84e-02b33ebba535)

![image](https://github.com/user-attachments/assets/956cb44b-649e-4c6c-a987-9628b106a327)

## logic逻辑
1. main function: excelConverter.py 主函数
2. readExcelClass.py: read excel and generate dataFrame 使用pandas读取excel文件
3. searchClass.py: read result and generate dataFrame 读取结果文件
4. exception_handler.py: catch the exception and emit signal 用于捕获异常并发送信号
5. BiDirectionalMap.py: create a BidirectionalMap used for search index and staff code 创建BidirectionalMap可以使用员工号快速搜索索引，也可以使用索引快速搜索员工号
6. bindClass.py(main logic): read the uploaded file and process the file, generate the result finally 读取上传的文件并处理，最终生成结果
7. excel2ppt_qt.py: "front end" logic, define the function when click the button and show error message 界面，按钮函数，展示异常

## run运行
dist/TaAMS_v1.2.2.exe
