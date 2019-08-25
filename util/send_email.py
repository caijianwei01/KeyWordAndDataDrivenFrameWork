#!/usr/bin/env python
# -*- coding:utf-8 -*-
import smtplib
from email.header import Header
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.utils import parseaddr, formataddr


# 格式化邮件地址
def _format_address(s):
    name, address = parseaddr(s)
    return formataddr((Header(name, 'utf-8').encode(), address))


# 添加附件
def _add_enclosure(file_path, msg):
    '''
    带附件的邮件可以看做包含若干部分的邮件：文本和各个附件本身，所以，可以构造一个MIMEMultipart对象代表邮件本身，
    然后往里面加上一个MIMEtext作为邮件正文，再继续往里面加上表示附件的MIMEBase对象即可
    :param file_path: 附件文件路径
    :param msg: 邮件对象
    :return:
    '''
    with open(file_path, 'rb') as f:
        # 设置附件的MIME和文件名，这里是png类型：
        mime = MIMEImage(f.read())
        # 加上必要的头信息
        mime.add_header('Content-Disposition', 'attachment', filename='test.png')
        mime.add_header('Content-ID', '0')
        # 添加到MIMEMultipart
        msg.attach(mime)


from_addr = '18367157420@163.com'
password = 'qwer128201209q@'
to_addr = '1031901787@qq.com'
smtp_server = 'smtp.163.com'

content = '''<html><body>
<h1>Hello</h1>
<p>send by <a href="http://www.python.org">Python</a></p>
<p><img src="cid:001"/></p>
</body></html>'''
# 邮件对象
msg = MIMEMultipart('mixed')
msg['From'] = _format_address('Python爱好者<%s>' % from_addr)
msg['To'] = _format_address('管理员<%s>' % to_addr)
msg['Subject'] = Header('来自SMTP的问候......', 'utf-8').encode()



# 添加图片
file_path_1 = r'C:\Users\18367\Desktop\001.png'
with open(file_path_1, 'rb') as f:
    image = MIMEImage(f.read())
    image.add_header('Content-ID', '001')
    image.add_header('Content-Disposition', 'attachment', filename='001.png')
    msg.attach(image)

msg.attach(MIMEText(content, 'html', 'utf-8'))

file_path = r'C:\Users\18367\Desktop\sanye.png'
# 添加图片附件
_add_enclosure(file_path, msg)

# SMTP协议默认端口是25
server = smtplib.SMTP(smtp_server, 25)
server.set_debuglevel(1)
server.login(from_addr, password)
server.sendmail(from_addr, [to_addr], msg.as_string())
server.quit()
