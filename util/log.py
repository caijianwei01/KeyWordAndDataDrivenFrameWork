#!/usr/bin/env python
# -*- coding: utf-8 -*-
import logging
import logging.config
from config.var_config import parent_dir_path

# 读取日志配置文件
log_cnf = parent_dir_path + '\\config\\Logger.conf'
logging.config.fileConfig(log_cnf)
# 选择一个日志格式
logger = logging.getLogger("example02")


def debug(message):
    # 定义debug级别日志打印方法
    logger.debug(message)


def info(message):
    # 定义info级别日志打印方法
    logger.info(message)


def warning(message):
    # 定义warning级别日志打印反法
    logger.warning(message)


if __name__ == '__main__':
    logger.debug('debug日志测试')
    logger.info('info日志测试')
    logger.warning('warning日志测试')
