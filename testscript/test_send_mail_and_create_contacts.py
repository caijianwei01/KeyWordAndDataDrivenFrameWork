#!/usr/bin/env python
# -*- coding:utf-8 -*-
from testscript.frame_work_fun import *


def test_send_mail_and_create_contacts():
    # 记录执行成功的测试用例个数
    successful_cases = 0
    # 记录需要执行的用例个数
    required_cases = 0
    # 根据Excel文件中的sheet名获取sheet对象
    case_sheet = excel_obj.get_sheet_by_name('测试用例')
    # 获取测试用例sheet中是否执行列对象
    is_execute_column = excel_obj.get_column(case_sheet, testCase_isExecute)
    try:
        # 循环遍历“测试用例”表中的测试用例
        for idx, col_value in enumerate(is_execute_column[1:]):
            # 用例sheet中第一行为标题行，无须执行
            # 循环遍历'测试用例'表中的测试用例，执行被设置为执行的用例
            if col_value.lower() == 'y':
                required_cases += 1
                # 获取"测试用例"表中第idx+2行数据,idx从0开始，第一行是标题行，第二行开始才是数据行
                case_row = excel_obj.get_row(case_sheet, idx + 2)
                # 获取第idx+2行"步骤sheet名"单元格内容，列表下标从零开始，所以需要减一
                case_step_sheet_name = case_row[testCase_testStepSheetName - 1]
                # 根据测试步骤名获取步骤sheet对象
                step_sheet = excel_obj.get_sheet_by_name(case_step_sheet_name)
                logging.info('----%s' % case_step_sheet_name)

                if case_row[testCase_frameWorkName - 1] == '数据':
                    logging.info('****** 调用数据驱动 ******')
                    # 获取数据驱动数据的sheet，数据驱动步骤的sheet
                    data_source_sheet = excel_obj.get_sheet_by_name(case_row[testCase_dataSourceSheetName - 1])
                    data_step_sheet = excel_obj.get_sheet_by_name(case_row[testCase_testStepSheetName - 1])
                    try:
                        # 通过数据驱动框架执行添加联系人
                        result = data_driver_fun(data_source_sheet, data_step_sheet)
                        if result:
                            logging.info('测试用例 “%s” 执行成功' % case_row[testCase_testCaseName - 1])
                            successful_cases += 1
                            write_test_result(case_sheet, idx + 2, 'testCase', 'pass')
                    except Exception:
                        logging.info('测试用例 “%s” 执行失败' % case_row[testCase_testCaseName - 1])
                        write_test_result(case_sheet, idx + 2, 'testCase', 'failed')
                elif case_row[testCase_frameWorkName - 1] == '关键字':
                    try:
                        logging.info('****** 调用关键字驱动 ******')
                        # 通过关键字驱动框架执行添加联系人
                        result = key_word_fun(case_row[testCase_testCaseName - 1], step_sheet)
                        if result:
                            logging.info('测试用例 “%s” 执行成功' % case_row[testCase_testCaseName - 1])
                            successful_cases += 1
                            write_test_result(case_sheet, idx + 2, 'testCase', 'pass')
                    except Exception:
                        logging.info('测试用例 “%s” 执行失败' % case_row[testCase_testCaseName - 1])
                        write_test_result(case_sheet, idx + 2, 'testCase', 'failed')
    except Exception:
        # 打印详细的异常堆栈信息
        logging.info(traceback.print_exc())
    finally:
        logging.info('共%d条用例，%d条需要被执行，本次执行通过%d条\n' % (len(is_execute_column) - 1,
                                                      required_cases, successful_cases))
        gl.get_value('driver').quit()


if __name__ == '__main__':
    test_send_mail_and_create_contacts()
