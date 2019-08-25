#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于实现将数据设置到剪贴板中
"""
import win32clipboard as w
import win32con


class Clipboard(object):
    """模拟Windows设置剪贴板"""

    # 读取剪贴板数据
    @staticmethod
    def get_text():
        # 打开剪贴板
        w.OpenClipboard()
        # 获取剪贴板中的数据
        data = w.GetClipboardData(win32con.CF_TEXT)
        # 关闭剪贴板
        w.CloseClipboard()
        # 返回剪贴板数据给调用者
        return data

    # 设置剪贴板内容
    @staticmethod
    def set_text(a_string):
        # 打开剪贴板
        w.OpenClipboard()
        # 清空剪贴板
        w.EmptyClipboard()
        # 将数据a_string写入剪贴板
        w.SetClipboardData(win32con.CF_UNICODETEXT, a_string)
        # 关闭剪贴板
        w.CloseClipboard()
