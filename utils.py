# coding=gbk
'''
Created on 2016��5��27��

@author: ����
'''
from ftplib import FTP
import logging
import win32com.client

def ftpFile(src, dest, Dfile, host, port, username, password):
    ftp=FTP() #���ñ���
    ftp.set_debuglevel(0) #�򿪵��Լ���2����ʾ��ϸ��Ϣ
    ftp.connect(host , int(port), timeout=30) #���ӵ�ftp sever�Ͷ˿�
    ftp.login(username ,password)#���ӵ��û���������
    
    ftp.cwd(dest) #����Զ��Ŀ¼
    ftp.set_pasv(True)
    ftp.storbinary("STOR " + Dfile, open(src, 'rb')) #�ϴ�Ŀ���ļ�
    ftp.set_debuglevel(0) #�رյ���ģʽ
    ftp.quit() #�˳�ftp


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