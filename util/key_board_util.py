#!/usr/bin/env python
# -*- coding:utf-8 -*-
"""
用于实现模拟键盘单个或组合按键
"""
import win32api
import win32con


class KeyBoardKeys(object):
    """模拟键盘按键类"""
    VK_CODE = {
        'enter': 13,
        'tab': 9,
        'ctrl': 17,
        'v': 86
    }

    @staticmethod
    def key_down(key_name):
        # 按下按键
        key_name = key_name.lower()
        win32api.keybd_event(KeyBoardKeys.VK_CODE[key_name], 0, 0, 0)

    @staticmethod
    def key_up(key_name):
        # 释放按键
        key_name = key_name.lower()
        win32api.keybd_event(KeyBoardKeys.VK_CODE[key_name], 0, win32con.KEYEVENTF_KEYUP, 0)

    @staticmethod
    def one_key(key):
        # 模拟单个按键
        KeyBoardKeys.key_down(key)
        KeyBoardKeys.key_up(key)

    @staticmethod
    def two_keys(key1, key2):
        # 模拟两个组合键，先按后释放
        KeyBoardKeys.key_down(key1)
        KeyBoardKeys.key_down(key2)
        KeyBoardKeys.key_up(key2)
        KeyBoardKeys.key_up(key1)
