# coding=gbk
'''
Created on 2016年1月23日

@author: 大雄
'''

import configparser
import os
import sys

config = None
def configure(file): 
    path = getHome()
        
    if not os.path.exists(path + "/conf/" + file):
        raise FileNotFoundError("conf file not exists: " + path + "/conf/" + file)
    
    c = configparser.ConfigParser()
    c.read(path + "/conf/" + file, "gbk")
    global config
    config = c["DEFAULT"]
    
def getHome():
    path = sys.path[0]
    if not os.path.isdir(path):
        path = os.path.split(os.path.realpath(__file__))[0]
    
    return path
        
