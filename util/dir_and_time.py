#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""用于获取当前日期及时间， 以及创建异常截图存放目录"""
import time
import os
from datetime import datetime
from config.var_config import screen_picture_dir


# 获取当前的日期
def get_current_date():
    time_tup = time.localtime()
    current_date = str(time_tup.tm_year) + '-' + str(time_tup.tm_mon) + '-' + str(time_tup.tm_mday)
    return current_date


# 获取当前的时间
def get_current_time():
    date_time = datetime.now()
    now_time = date_time.strftime('%H-%M-%S-%f')
    return now_time


# 创建截图存放的目录
def create_current_date_dir():
    dir_name = os.path.join(screen_picture_dir, get_current_date())
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    return dir_name


if __name__ == '__main__':
    print(get_current_date())
    print(get_current_time())
    print(create_current_date_dir())
