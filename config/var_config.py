#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于定义整个框架中所需要的一些全局常量值
"""
import os

# 谷歌浏览器驱动器路径
chrome_driver_file_path = 'D:\zdh\chromedriver.exe'

# 获取当前文件所在目录的父目录的绝对路径
parent_dir_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 异常截图保存路径，\\ 双斜杠是为了转义\
screen_picture_dir = parent_dir_path + '\\exception_pictures\\'
# 测试数据文件存放绝对路径
data_file_path = parent_dir_path + '\\testdata\\163邮箱发送邮件.xlsx'

# 测试数据文件中，测试用例表中部分列对应的数字序号
testCase_testCaseName = 2
testCase_testCaseDescribe = 3
testCase_frameWorkName = 4
testCase_testStepSheetName = 5
testCase_dataSourceSheetName = 6
testCase_isExecute = 7
testCase_runTime = 8
testCase_testResult = 9

# 用例步骤表中，部分列对应的数字序号
testStep_testStepDescribe = 2
testStep_keyWord = 3
testStep_locationType = 4
testStep_locatorExp = 5
testStep_operateValue = 6
testStep_runTime = 7
testStep_testResult = 8
testStep_errorInfo = 9
testStep_errorPic = 10

# 数据源表中，部分列对应的序号
dataSource_email = 2
dataSource_isExecute = 6
dataSource_runTime = 7
dataSource_result = 8

if __name__ == '__main__':
    print(parent_dir_path)
    print(screen_picture_dir)
    print(data_file_path)