# coding=gbk
'''
Created on 2016年5月27日

@author: 大雄
'''
from ftplib import FTP
import logging
import win32com.client

def ftpFile(src, dest, Dfile, host, port, username, password):
    ftp=FTP() #设置变量
    ftp.set_debuglevel(0) #打开调试级别2，显示详细信息
    ftp.connect(host , int(port), timeout=30) #连接的ftp sever和端口
    ftp.login(username ,password)#连接的用户名，密码
    
    ftp.cwd(dest) #更改远程目录
    ftp.set_pasv(True)
    ftp.storbinary("STOR " + Dfile, open(src, 'rb')) #上传目标文件
    ftp.set_debuglevel(0) #关闭调试模式
    ftp.quit() #退出ftp


def CheckProcExistByPN(progname):
    try:
        WMI = win32com.client.GetObject('winmgmts:') 
        processCodeCov = WMI.ExecQuery('select * from Win32_Process where Name="%s"' % progname)
    except Exception as e:
        logging.debug(progname + " error : " + str(e))
        return 0
    if len(processCodeCov) > 0:
        logging.debug(progname + " exist")
        return 1
    else:
        logging.debug(progname + " not exist")
        return 0

if __name__ == '__main__':
    import os
    logging.basicConfig(level=logging.DEBUG,
                format='%(asctime)s [%(levelname)s] %(filename)s[line:%(lineno)d] %(message)s',
                datefmt='%a, %d %b %Y %H:%M:%S')
    if CheckProcExistByPN("DataCapture.exe") != 1:
        os.system("C:\Python34\python.exe")