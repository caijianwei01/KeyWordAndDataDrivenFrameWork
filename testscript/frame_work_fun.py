#!/usr/bin/env python
# -*- coding:utf-8 -*-
from action.page_action import *
from util.log import *
from config.var_config import *
import traceback
import time

# 获取全局变量excel_obj
excel_obj = gl.get_value('excel_obj')


def data_driver_fun(data_sheet, step_sheet):
    """数据驱动"""
    # 获取数据表中是否执行列对象
    is_execute_column = excel_obj.get_column(data_sheet, dataSource_isExecute)
    # 获取数据源表中"电子邮件"列对象
    email_column = excel_obj.get_column(data_sheet, dataSource_email)
    # 获取测试步骤表中存在数据区域的行数
    step_rows_num = excel_obj.get_rows_num(step_sheet)
    # 记录成功执行的数据条数
    success_datas = 0
    # 记录被设置为执行的数据条数
    required_datas = 0
    try:
        # 遍历“联系人”表中的数据，再通过遍历“创建联系人”表中的步骤，添加联系人
        for idx, col_value in enumerate(is_execute_column[1:]):
            # idx下标从0开始
            # 遍历数据源表，准备进行数据驱动测试，因为第一行是标题行，所以从第二行开始遍历
            if col_value == 'y':
                logging.info('开始添加联系人 “%s”' % email_column[idx + 1])
                required_datas += 1
                # 定义记录执行成功步骤数变量
                success_steps = 0
                # 遍历“创建联系人”表中的步骤
                for index in range(2, step_rows_num + 1):
                    # 获取数据驱动测试步骤表中第index行对象
                    rows_value = excel_obj.get_row(step_sheet, index)
                    # 获取当前行中的关键字、定位方式、定位表达式式、操作值
                    key_word = rows_value[testStep_keyWord - 1]
                    location_type = rows_value[testStep_locationType - 1]
                    locator_exp = rows_value[testStep_locatorExp - 1]
                    operate_value = rows_value[testStep_operateValue - 1]

                    # 将操作值为数字类型的数据转成字符串类型，方便字符串拼接
                    if isinstance(operate_value, int):
                        operate_value = str(operate_value)
                    if operate_value != 'None' and operate_value.isalpha():
                        # 如果operate_value不为空且为字母字符串，从数据源表中根据坐标获取对应单元格的数据
                        coordinate = operate_value + str(idx + 2)
                        operate_value = excel_obj.get_cell_of_value(data_sheet, coordinate=coordinate)

                    # 构造需要执行的python语句，
                    # 对应的是page_action.py文件中的页面动作函数调用的字符串表示
                    tmp_str = "'%s', '%s'" % (location_type.lower(), locator_exp) \
                        if location_type != 'None' and locator_exp != 'None' else ""
                    if tmp_str:
                        tmp_str += ", '" + operate_value + "'" if operate_value != 'None' \
                                                                  and operate_value != '是' else ""
                    else:
                        tmp_str += "'" + operate_value + "'" if operate_value != 'None' else ""
                    run_str = key_word + "( " + tmp_str + ")"
                    try:
                        # 通过eval函数，将拼接的页面动作函数调用的字符串表示当成有效的python表达式执行，
                        # 从而执行测试步骤的sheet中关键字在page_action.py文件中对应的映射方法，来完成对页面元素的操作。
                        if operate_value != '否':
                            # 当operate_value值为“是”时，表示单击星标联系人复选框
                            eval(run_str)
                    except Exception:
                        logging.info('执行步骤“%s”发生异常: %s' % (rows_value[testStep_testStepDescribe - 1],
                                                           traceback.print_exc()))
                        write_test_result(data_sheet, idx + 2, 'dataSheet', 'failed')
                        break
                    else:
                        success_steps += 1
                        logging.info('执行步骤“%s”成功' % rows_value[testStep_testStepDescribe - 1])
                    time.sleep(0.5)
                if step_rows_num == success_steps + 1:
                    success_datas += 1
                    # 如果成功执行的步骤数等于步骤表中给出的步骤数
                    # 说明第idx+2行的数据执行通过，写入通过信息
                    write_test_result(data_sheet, idx + 2, 'dataSheet', 'pass')
                else:
                    # 写入失败信息
                    write_test_result(data_sheet, idx + 2, 'dataSheet', 'failed')
                time.sleep(0.1)
            else:
                # 将不需要执行的数据行的执行时间和执行结果单元格清空
                write_test_result(data_sheet, idx + 2, 'dataSheet', '')
        if required_datas == success_datas:
            # 只有当成功执行的数据条数等于被设置为需要执行的数据条数时，才表示调用数据驱动的测试用例执行通过
            return 1
        return 0
    except Exception as ex:
        raise ex


def key_word_fun(case_name, step_sheet):
    """关键字驱动"""
    # 获取"步骤sheet名"中的步骤数
    step_num = excel_obj.get_rows_num(step_sheet)
    # 记录测试用例col_value的步骤成功数
    successful_steps = 0
    logging.info('开始执行用例 "%s"，用例共%s步' % (case_name, step_num))
    try:
        for step in range(2, step_num + 1):
            # 第一行为标题行，无须执行，所以从2开始
            # 获取步骤sheet中第step行对象
            step_row = excel_obj.get_row(step_sheet, step)
            # 获取关键字、定位方式、定位表达式、操作值
            key_word = step_row[testStep_keyWord - 1]
            location_type = step_row[testStep_locationType - 1]
            locator_exp = step_row[testStep_locatorExp - 1]
            operate_value = step_row[testStep_operateValue - 1]

            # 将操作值为数字类型的数据转成字符串类型，方便字符串拼接
            if isinstance(operate_value, int):
                operate_value = str(operate_value)

            expression_str = ''
            # 构造需要执行的python语句，
            # 对应的是page_action.py文件中的页面动作函数调用的字符串表示
            if key_word and operate_value != 'None' and location_type == 'None' and locator_exp == 'None':
                expression_str = key_word + "('" + operate_value + "')"
            elif key_word and operate_value == 'None' and location_type == 'None' and locator_exp == 'None':
                expression_str = key_word + "()"
            elif key_word and location_type and operate_value and locator_exp == 'None':
                expression_str = key_word + "('" + location_type + "', '" + operate_value + "')"
            elif key_word and location_type and locator_exp and operate_value != 'None':
                expression_str = key_word + "('" + location_type + "', '" + \
                                 locator_exp + "', '" + operate_value + "')"
            elif key_word and location_type and locator_exp and operate_value == 'None':
                expression_str = key_word + "('" + location_type + "', '" + locator_exp + "')"
            try:
                # 在测试执行时间列写入执行时间
                excel_obj.write_cell_current_time(step_sheet, row_no=step, col_no=testStep_runTime)
                # 通过eval函数，将拼接的页面动作函数调用的字符串表示当成有效的python表达式执行，
                # 从而执行测试步骤的sheet中关键字在page_action.py文件中对象的映射方法，来完成对页面元素的操作。
                eval(expression_str)
            except Exception:
                # 获取异常屏幕截图
                capture_pic = capture_screen()
                # 获取详细的异常堆栈信息
                error_info = traceback.format_exc()
                # 在测试步骤Sheet中写入失败信息
                write_test_result(step_sheet, step, "caseStep", "failed", error_info, capture_pic)
                logging.info('案例：%s, 步骤: %s, 执行失败！' % (case_name,
                                                       step_row[testStep_testStepDescribe - 1]))
                break
            else:
                # 在测试步骤Sheet中写入成功信息
                write_test_result(step_sheet, step, 'caseStep', 'pass')
                # 每成功一步，successful_steps变量自增1
                successful_steps += 1
                logging.info('步骤: %s, 执行成功！' % step_row[testStep_testStepDescribe - 1])
                time.sleep(0.5)
        if successful_steps == step_num - 1:
            # 当测试用例步骤中所有的步骤都执行成功，才认为此测试用例执行通过
            return 1
        return 0
    except Exception as ex:
        raise ex


def write_test_result(sheet, row_no, col_no, test_result, error_info=None, pic_path=None):
    """用例或用例步骤执行结束后，向Excel中写入执行结果信息"""
    # 测试通过结果信息为绿色，失败为红色
    color_dict = {'pass': 'green', 'failed': 'red', '': None}
    # 因为“测试用例”工作表和“用例步骤sheet表”中都有测试执行时间和
    # 测试结果列，定义此字典对象是为了区分具体应该写哪个工作表。
    cols_dict = {
        'testCase': [testCase_runTime, testCase_testResult],
        'caseStep': [testStep_runTime, testStep_testResult],
        'dataSheet': [dataSource_runTime, dataSource_result]}
    try:
        # 在测试步骤sheet中，写入测试结果
        excel_obj.write_cell(sheet, content=test_result, row_no=row_no, col_no=cols_dict[col_no][1],
                             style=color_dict[test_result])
        time.sleep(0.1)
        if test_result == '':
            # 清空时间单元格内容
            excel_obj.write_cell(sheet, content='', row_no=row_no, col_no=cols_dict[col_no][0])
            time.sleep(0.1)
        else:
            # 在测试步骤sheet中，写入测试时间
            excel_obj.write_cell_current_time(sheet, row_no=row_no, col_no=cols_dict[col_no][0])
            time.sleep(0.1)
        if error_info and pic_path:
            # 在测试步骤sheet中，写入异常信息
            excel_obj.write_cell(sheet, content=error_info, row_no=row_no, col_no=testStep_errorInfo)
            time.sleep(0.1)
            # 在测试步骤sheet中，写入异常截图路径
            excel_obj.write_cell(sheet, content=pic_path, row_no=row_no, col_no=testStep_errorPic)
        else:
            if col_no == 'caseStep':
                # 在测试步骤sheet中，清空错误信息单元格
                excel_obj.write_cell(sheet, content='', row_no=row_no, col_no=testStep_errorInfo)
                time.sleep(0.1)
                # 在测试步骤sheet中，清空错误截图单元格
                excel_obj.write_cell(sheet, content='', row_no=row_no, col_no=testStep_errorPic)
    except Exception:
        logging.debug('写入Excel文件出错: %s' % traceback.print_exc())
