#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于实现定位页面元素
"""
from selenium.webdriver.support.ui import WebDriverWait


def get_element(driver, location_type, locator_exp):
    """
    获取单个页面元素对象
    :param driver:
    :param location_type: 定位类型
    :param locator_exp: 定位表达式
    :return:返回找到的页面元素对象
    """
    try:
        # 显示等待，默认是每0.5秒找一次元素
        element = WebDriverWait(driver, 30).until(lambda x: x.find_element(by=location_type, value=locator_exp))
        return element
    except Exception as ex:
        raise ex


# 获取相同页面多个相同元素对象，以list返回
def get_elements(driver, location_type, locator_exp):
    try:
        elements = WebDriverWait(driver, 30).until(lambda x: x.find_elements(by=location_type, value=locator_exp))
        return elements
    except Exception as ex:
        raise ex
