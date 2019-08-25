#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于实现智能等待页面元素的出现
"""
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class WaitUtil(object):

    def __init__(self, driver):
        self.location_type_dict = {
            'id': By.ID,
            'xpath': By.XPATH,
            'name': By.NAME,
            'class_name': By.CLASS_NAME,
            'tag_name': By.TAG_NAME,
            'link_text': By.LINK_TEXT,
            'partial_link_text': By.PARTIAL_LINK_TEXT
        }
        self.driver = driver
        self.wait = WebDriverWait(self.driver, 30)

    def presence_of_element_located(self, location_type, locator_exp):
        """
        显示等待页面元素出现在DOM树中，但不一定可见，存在则返回该页面元素对象
        """
        location_type = location_type.lower()
        try:
            if location_type in self.location_type_dict:
                element = self.wait.until(
                    EC.presence_of_element_located((
                        self.location_type_dict[location_type], locator_exp)))
                return element
            else:
                raise TypeError('未找到定位方式：%r，请确认定位方法是否写正确' % locator_exp)
        except Exception as ex:
            raise ex

    def frame_available_and_switch_to_it(self, location_type, locator_exp):
        """
        :param location_type: 定位类型
        :param locator_exp: 定位表达式
        :return:
        检查frame是否存在，存在则切换进frame控件中
        """
        try:
            self.wait.until(EC.frame_to_be_available_and_switch_to_it(
                (self.location_type_dict[location_type.lower()], locator_exp)))
        except Exception as ex:
            raise ex

    def visibility_element_located(self, location_type, locator_exp):
        """显示等待页面元素的出现"""
        try:
            element = self.wait.until(
                EC.visibility_of_element_located((self.location_type_dict[location_type.lower()], locator_exp)))
            return element
        except Exception as ex:
            raise ex


if __name__ == '__main__':
    from selenium import webdriver
    import time

    i_driver = webdriver.Chrome(executable_path='D:\zdh\chromedriver.exe')
    i_driver.get('https://mail.126.com')
    wait_util = WaitUtil(i_driver)
    wait_util.frame_available_and_switch_to_it('xpath', '//iframe[starts-with(@id, "x-URS-iframe")]')
    e = wait_util.visibility_element_located('xpath', '//input[@name="email"]')
    e.send_keys('success')
    time.sleep(5)
    ele = wait_util.presence_of_element_located('xpath', '//input[@name="email"]')
    print(ele)
    i_driver.quit()
