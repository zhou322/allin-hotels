# -*- coding: utf-8 -*-

import os
import poplib
import email
import time
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr


class QQEmail(object):
    def __init__(self, user, password):
        self.User = user  # 邮箱用户名
        self.Pass = password  # 邮箱密码
        self.server = None  # 初始化server，调用login_email方法后更新server
        self.mails = None  # 初始化邮箱信息列表，调用get_email_lists方法后更新列表

    def login_email(self):
        # 登录邮箱
        pop3_server = 'imap.exmail.qq.com'
        try:
            server = poplib.POP3(pop3_server, 110, timeout=50)

            # 身份认证:
            server.user(self.User)
            server.pass_(self.Pass)
            self.server = server
            return {"status": True, 'info': f'{self.User} login successful!'}
        except BaseException as e:
            return {"status": False, 'info': f'Login failed: {e}'}

    # 字符编码转换
    def decode_str(self, str_in):
        try:
            value, charset = decode_header(str_in)[0]
            if charset:
                value = value.decode(charset)
            return value
        except:
            return str_in

    def guess_charset(self, msg):
        charset = msg.get_charset()
        if charset is None:
            content_type = msg.get('Content-Type', '').lower()
            pos = content_type.find('charset=')
            if pos >= 0:
                charset = content_type[pos + 8:].strip()
        return charset

    # indent用于缩进显示:
    def print_info(self, msg, indent=0):
        if indent == 0:
            for header in ['From', 'To', 'Subject']:
                value = msg.get(header, '')
                if value:
                    if header == 'Subject':
                        value = self.decode_str(value)
                    else:
                        hdr, addr = parseaddr(value)
                        name = self.decode_str(hdr)
                        value = u'%s <%s>' % (name, addr)
                print('%s%s: %s' % ('  ' * indent, header, value))
        if (msg.is_multipart()):
            parts = msg.get_payload()
            for n, part in enumerate(parts):
                print('%spart %s' % ('  ' * indent, n))
                print('%s--------------------' % ('  ' * indent))
                self.print_info(part, indent + 1)
        else:
            content_type = msg.get_content_type()
            if content_type == 'text/plain' or content_type == 'text/html':
                content = msg.get_payload(decode=True)
                charset = self.guess_charset(msg)
                if charset:
                    content = content.decode(charset)
                print('%sText: %s' % ('  ' * indent, content + '...'))
            else:
                print('%sAttachment: %s' % ('  ' * indent, content_type))

    # 获取邮箱信息列表
    def get_email_lists(self):
        try:
            resp, mails, octets = self.server.list()  # list()返回所有邮件的编号:
            self.mails = mails
            return {"status": True, 'info': f'成功获取邮件列表!'}
        except BaseException as e:
            return {"status": False, 'info': f'邮件列表获取失败: {e}'}

    # 解析邮件,获取附件
    def get_att(self, msg_in):
        attachment_files = []
        i = 1
        for part in msg_in.walk():
            # 获取附件名称类型
            file_name = part.get_filename()
            # contType = part.get_content_type()
            if file_name:
                h = email.header.Header(file_name)

                # 对附件名称进行解码
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # 将附件名称可读化
                    filename = self.decode_str(str(filename, dh[0][1]))
                    # print(filename)
                    # filename = filename.encode("utf-8")

                # 下载附件
                data = part.get_payload(decode=True)
                path = r"附件" # 在指定目录下创建文件,如果不存在则创建目录
                if not os.path.exists(path):
                    os.makedirs(path)
                att_file = open(path + '\\' + filename, 'wb') # 注意二进制文件需要用wb模式打开
                attachment_files.append(filename)
                att_file.write(data)  # 保存附件
                att_file.close()

                print(f'附件({i}): {filename}')
                i += 1
        return attachment_files

    # 解析邮件
    def parser_mail(self, index):
        '''
        :param index: 邮件索引
        :return: 邮件正文、时间、主题、发件人的字典
        '''

        # 1、获取邮件原文
        try:
            resp, lines, octets = self.server.retr(index)  # 获取第index封邮件，lines存储了邮件的原始文本的每一行
        except:
            try:  # 如果获取邮件失败，尝试重新登录邮箱再获取
                self.login_email()
                self.get_email_lists()
                resp, lines, octets = self.server.retr(index)
            except:  # 如果还是失败，返回False
                return False

        # 2、拼接邮件
        try:
            msg_content = b'\n'.join(lines).decode('gbk')  # 邮件的原始文本
        except:
            try:
                msg_content = b'\n'.join(lines).decode('utf-8')  # 邮件的原始文本
            except:
                return False

        # 3、解析邮件内容
        try:
            msg = Parser().parsestr(msg_content)
        except:
            return False

        # 4、解析邮件主题(标题)
        try:
            Subject = self.decode_str(msg.get("Subject"))
        except BaseException as e:
            return False

        # 5、解析邮件时间
        try:
            Date = time.strptime(self.decode_str(msg.get("Date"))[0:24], '%a, %d %b %Y %H:%M:%S')
            Date = time.mktime(Date)  # 获取邮件的接收时间,格式化收件时间
        except:
            return False

        # 6、解析发件人
        try:
            From = self.decode_str(msg.get("From")).split(' ')[-1]
        except:
            From = '<None>'

        return {
            'From': From,
            'Date': Date,
            'Subject': Subject,
            'Msg': msg,
        }

    # 退出server
    def server_quit(self):
        self.server.quit()


if __name__ == "__main__":

    Q = QQEmail(user='邮箱', password='密码')  # 初始化类
    login_info = Q.login_email()  # 登录邮箱
    print(login_info['info'])

    if login_info['status']:
        email_lists = Q.get_email_lists()  # 获取邮件列表
        if email_lists['status']:
            indexs = range(len(Q.mails), 0, -1)[-10:]  # 获取最近的10封邮件索引

            # 从最近的邮件开始，依次遍历所有邮件
            for index in indexs:
                mail_msg = Q.parser_mail(index)  # 解析邮件
                if mail_msg:
                    # Q.print_info(mail_msg['Msg'])  # 输入邮件内容
                    Q.get_att(mail_msg['Msg'])  # 下载邮件中的附件
