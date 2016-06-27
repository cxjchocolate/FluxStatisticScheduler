# coding=gbk
'''
Created on 2016年5月4日

@author: 大雄
'''

import pyodbc
import csv
import logging
import env
from apscheduler.scheduler import Scheduler
from utils import ftpFile, CheckProcExistByPN
import subprocess
import os
import time
import datetime


def getFluxDetailsByStoreID(db, storeID, status ):
    
    try:
        '''
        connects with odbc
        '''        
        conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db + ";")
        cur = conn.cursor()
        
        '''
        add storeID into results where data not uploaded(FUpload=0)
        '''            
        strsql = r"select ID, DeviceID, Date, FOut, FIn, FAll, FUpload, '{0}' as storeID from FluxDetails where FUpload={1}".format(storeID, status)
        cur.execute(strsql)
        logging.debug(strsql)
        t = list(cur)
        return t
    except Exception as e:
        logging.debug(e)
    finally:
        conn.close()
        
    return None
    
def writeListToCSV(csv_file, data_list ):
    try:
        with open(csv_file, 'w') as csvfile:
            writer = csv.writer(csvfile, dialect='excel', quoting=csv.QUOTE_ALL)
            ids = []
            for data in data_list:
                ids.append(data.ID)
                writer.writerow(data)
            return ids
    except Exception as e:
            logging.info("I/O error({0})".format(e))    
    return None

def updateFluxDetailsByStoreID(db, storeID, ids, status):
    
    try:
        '''
        connects with odbc
        '''        
        conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db + ";")
        cur = conn.cursor()
        
        '''
        add storeID into results
        '''
        idstr = ','.join(str(i) for i in ids)
        strsql = r"update FluxDetails set FUpload={0} where id in ({1})".format(status, idstr)
        logging.debug(strsql)
        result = cur.execute(strsql).rowcount
        conn.commit()
        return result
    except Exception as e:
        logging.debug(e)
    finally:
        conn.rollback()
        conn.close()
        
    return None

def transferFluxDetailsByStatus(db, storeID, dest, host, port, username, password, status):
    try:
        csv = env.config.get("TEMPDIR") + '/FluxDetails.' + env.config.get("STOREID") + "." + time.strftime("%Y%m%d%H%M%S", time.localtime())  + ".csv"
        '''
        export data from access db where status = %status
        '''
        data2 = getFluxDetailsByStoreID(db, storeID, status)
        if (not data2 or len(data2) == 0):
            logging.info("no new data is found")
            return
        '''
        write data to csv file
        '''   
        ids = writeListToCSV(csv, data2)
        '''
        update status to 5
        '''
        updateFluxDetailsByStoreID(db, storeID, ids, 5)
        '''
        ftp csv file to server
        '''
        ftpFile(csv, dest, os.path.basename(csv), host, port, username, password)
        '''
        update status to 1
        '''        
        updateFluxDetailsByStoreID(db, storeID, ids, 1)
        '''
        delete temp file
        '''
        os.remove(csv)
    except Exception as e:
        logging.debug(e)


def getHistDetailsByStoreID(db, storeID, beforeDate ):
    
    try:
        '''
        connects with odbc
        '''        
        conn = pyodbc.connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + db + ";")
        cur = conn.cursor()
        
        '''
        add storeID into results where data not uploaded(FUpload=0)
        '''
        str_beforeDate = beforeDate.strftime("%Y-%m-%d %H:%M:%S")      
        strsql = r"select ID, DeviceID, Date, FOut, FIn, FAll, FUpload, '{0}' as storeID from FluxDetails where Date < cdate('{1}') and FUpload=0".format(storeID, str_beforeDate)
        logging.debug(strsql)
        cur.execute(strsql)
        t = list(cur)
        return t
    except Exception as e:
        logging.debug(e)
    finally:
        conn.close()
        
    return None

def transferHistDetails(db, storeID, dest, host, port, username, password):
    try:
        csv = env.config.get("TEMPDIR") + '/FluxDetails.' + env.config.get("STOREID") + "." + time.strftime("%Y%m%d%H%M%S", time.localtime())  + ".csv"
        
        '''
        export data from access db
        '''
        data2 = getHistDetailsByStoreID(db, storeID, datetime.date.today())
        if (not data2 or len(data2) == 0):
            logging.info("no new data is found")
            return
        
        '''
        write data to csv file
        '''   
        ids = writeListToCSV(csv, data2)
        '''
        update status to 5
        '''
        updateFluxDetailsByStoreID(db, storeID, ids, 5)
        '''
        ftp csv file to server
        '''
        ftpFile(csv, dest, os.path.basename(csv), host, port, username, password)
        
        updateFluxDetailsByStoreID(db, storeID, ids, 1)
        '''
        delete temp file
        '''
        os.remove(csv)
    except Exception as e:
        logging.debug(e)

           
def sync(db, storeID, progname, prog, dest, host, port, username, password):
    watchDog(progname, prog)
    ftpFile("status.dat", dest, storeID + ".dat", host, port, username, password)
    
    transferFluxDetailsByStatus(db, storeID, 
                                 dest, host, port, username, password, 5)


def watchDog(progname, prog):
    '''
    DataCapture.exe
    '''
    try:
        if CheckProcExistByPN(progname) != 1:
            logging.info("starting " + progname + ": " + prog)
            subprocess.Popen(prog);
    except Exception as e:
        logging.debug(e)
    

def main(standalone=True):
    
    env.configure("conf.ini")
    
    '''
    transfer History Detail for the first time
    when program is executed
    '''
    if env.config.get("transfer_hist_detail_on_loading") and env.config.get("transfer_hist_detail_on_loading") == "1":
        transferHistDetails(env.config.get("DB_FILE"), env.config.get("STOREID"), 
                         env.config.get("FTP_HOME"), env.config.get("FTP_HOST"),
                         env.config.get("FTP_PORT"), env.config.get("FTP_USERNAME"),
                         env.config.get("FTP_PASSWORD"))
     
    '''
    create the scheduler for two job:
    1. check the client's alive
    2. transfer details to server
    '''
    scheduler = Scheduler(standalone=standalone)
      
    scheduler.add_interval_job(sync, 
                               minutes=int(env.config.get("sync_interval")), 
                               args=(env.config.get("DB_FILE"), env.config.get("STOREID"),
                                     env.config.get("PROGNAME"), env.config.get("PROG"),
                                     env.config.get("FTP_HOME"), env.config.get("FTP_HOST"),env.config.get("FTP_PORT"), 
                                     env.config.get("FTP_USERNAME"),env.config.get("FTP_PASSWORD")))
       
    scheduler.add_cron_job(transferFluxDetailsByStatus, 
                           day_of_week=env.config.get("upload_day_of_week"), hour=env.config.get("upload_hour"), minute=env.config.get("upload_minute"), 
                           args=(env.config.get("DB_FILE"), env.config.get("STOREID"),
                                 env.config.get("FTP_HOME"), env.config.get("FTP_HOST"),env.config.get("FTP_PORT"), 
                                 env.config.get("FTP_USERNAME"),env.config.get("FTP_PASSWORD"), 
                                 0))
    scheduler.start()


if __name__ == '__main__':
    main()
    
    