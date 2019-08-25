#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于实现具体的页面动作，比如在输入框中输入数据，单击页面按钮等
"""
from selenium import webdriver
from config.var_config import chrome_driver_file_path
from util.object_map import get_element
from util.clipboard_util import Clipboard
from util.key_board_util import KeyBoardKeys
from util.dir_and_time import *
from util.wait_util import WaitUtil
from selenium.webdriver.chrome.options import Options
import util.global_var as gl
import time

# 定义全局driver变量
driver = None
# 全局的等待类实例对象
wait_util = None


def open_browser(browser_name):
    # 打开浏览器
    global driver, wait_util
    try:
        if browser_name.lower() == 'ie':
            pass
        elif browser_name.lower() == 'firefox':
            pass
        elif browser_name.lower() == 'chrome':
            # 创建Chrome浏览器的一个Options实例对象
            chrome_options = Options()
            # 向Options实例中添加禁用扩展插件的设置参数项
            chrome_options.add_argument('--disable-extensions')
            # 添加屏蔽--ignore-certificate-errors提示信息的设置参数项
            chrome_options.add_experimental_option('excludeSwitches', ['ignore-certificate-errors'])
            # 添加浏览器最大化的设置参数项，已启动就最大化
            chrome_options.add_argument('--start-maximized')
            # 通过配置参数禁止data;的出现
            # chrome_options.add_argument("--user-data-dir=C:/Users/18367/AppData/Local/Google/Chrome/User Data/Default")
            # 启动带有自定义设置的Chrome浏览器
            driver = webdriver.Chrome(executable_path=chrome_driver_file_path, options=chrome_options)
            # 将driver设置成为跨模块的变量
            gl.set_value('driver', driver)
        # driver对象创建成功后，创建等待类实例对象
        wait_util = WaitUtil(driver)
    except Exception as ex:
        raise ex


def visit_url(url):
    # 访问某个网址
    global driver
    try:
        driver.get(url)
    except Exception as ex:
        raise ex


def close_browser():
    # 关闭浏览器
    global driver
    try:
        driver.quit()
    except Exception as ex:
        raise ex


def sleep(sleep_seconds):
    # 强制等待
    try:
        time.sleep(int(sleep_seconds))
    except Exception as ex:
        raise ex


def clear(location_type, locator_exp):
    # 清除输入框默认内容
    global driver
    try:
        get_element(driver, location_type, locator_exp).clear()
    except Exception as ex:
        raise ex


def input_string(location_type, locator_exp, input_content):
    # 在页面输入框中输入数据
    global driver
    try:
        # 先清空输入框，再输入内容
        clear(location_type, locator_exp)
        get_element(driver, location_type, locator_exp).send_keys(input_content)
    except Exception as ex:
        raise ex


def click(location_type, locator_exp):
    # 单击页面元素
    global driver
    try:
        get_element(driver, location_type, locator_exp).click()
    except Exception as ex:
        raise ex


def assert_string_in_page_source(assert_string, timeouts=20):
    # 断言页面源码中是否存在某关键字或关键字符串
    global driver
    try:
        for i in range(timeouts):
            if assert_string in driver.page_source:
                break
            time.sleep(0.5)
        assert assert_string in driver.page_source, '%s not found in page source!' % assert_string
    except AssertionError as ex:
        raise AssertionError(ex)
    except Exception as ex:
        raise ex


def assert_title(title_str):
    # 断言页面标题是否存在给定的关键字符串
    global driver
    try:
        assert title_str in driver.title, '%s not found in title!' % title_str
    except AssertionError as ex:
        raise AssertionError(ex)
    except Exception as ex:
        raise ex


def get_title():
    # 获取页面标题
    global driver
    try:
        return driver.title
    except Exception as ex:
        raise ex


def get_page_source():
    # 获取页面源代码
    global driver
    try:
        return driver.page_source
    except Exception as ex:
        raise ex


def switch_to_frame(location_type, locator_exp):
    # 切换进入frame
    global driver
    try:
        element = get_element(driver, location_type, locator_exp)
        driver.switch_to_frame(element)
    except Exception as ex:
        raise ex


def switch_to_default_content():
    # 切除frame
    global driver
    try:
        driver.switch_to_default_content()
    except Exception as ex:
        raise ex


def paste_string(paste_str):
    # 模拟Ctrl + V 操作
    try:
        Clipboard.set_text(paste_str)
        # 等待2秒，防止代码执行的太快，而未成功粘贴内容
        time.sleep(2)
        KeyBoardKeys.two_keys('ctrl', 'v')
    except Exception as ex:
        raise ex


def press_tab_key():
    # 模拟Tab键
    try:
        KeyBoardKeys.one_key('tab')
    except Exception as ex:
        raise ex


def press_enter_key():
    # 模拟Enter键
    try:
        KeyBoardKeys.one_key('enter')
    except Exception as ex:
        raise ex


def maximize_browser():
    # 窗口最大化
    global driver
    try:
        driver.maximize_window()
    except Exception as ex:
        raise ex


def capture_screen():
    # 截取屏幕图片，\\是为了转义反斜杠
    global driver
    current_time = get_current_time()
    pic_name_and_path = str(create_current_date_dir()) + '\\' + str(current_time) + '.png'
    try:
        driver.get_screenshot_as_file(pic_name_and_path.replace('\\', r'\\'))
    except Exception as ex:
        raise ex
    else:
        return pic_name_and_path


def wait_presence_of_element_located(location_type, locator_exp):
    # 显式等待页面元素出现在DOM中，但并不一定可见，存在则返回该页面元素对象
    global wait_util
    try:
        wait_util.presence_of_element_located(location_type, locator_exp)
    except Exception as ex:
        raise ex


def wait_frame_available_and_switch_to_it(location_type, locator_exp):
    # 检查frame是否存在，存在则切换进frame控件中
    global wait_util
    try:
        wait_util.frame_available_and_switch_to_it(location_type, locator_exp)
    except Exception as ex:
        raise ex


def wait_visibility_of_element_located(location_type, locator_exp):
    # 显式等待页面元素出现在DOM中，并且可见， 存在返回该页面元素对象
    global wait_util
    try:
        wait_util.visibility_element_located(location_type, locator_exp)
    except Exception as ex:
        raise ex
