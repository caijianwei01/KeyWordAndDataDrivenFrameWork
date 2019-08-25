#!/usr/bin/env python
# encoding: utf-8
from util.parse_excel import ParseExcel
from config.var_config import data_file_path
import util.global_var as gl


# 创建解析Excel对象
excel_obj = ParseExcel()
# 将Excel数据文件加载到内存
excel_obj.load_workbook(data_file_path)
# 设置全局变量
gl.set_value('excel_obj', excel_obj)
