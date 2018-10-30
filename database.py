#! /usr/bin/env python
# -*- coding:utf-8 -*-
# vim:fenc=utf-8
#
# Copyright © 2014 jay <hujiangyi@dvt.dvt.com>
#

import ConfigParser
from pymongo import MongoClient

class database :
    def __init__(self):
        databaseIp = self.getConfig('database','ip')
        databasePort = self.getConfig('database','port')
        databaseUsername = self.getConfig('database','username')
        databasePassword = self.getConfig('database','password')
        conn = MongoClient(databaseIp, int(databasePort))
        db = conn.cm30
        db.authenticate(databaseUsername, databasePassword)
        self.cmlog = db.cmlog

    def isMacExist(self,mac):
        result = self.cmlog.find_one({"mac":mac,"isFail":False})
        return not result is None

    def isSnExist(self,sn):
        result = self.cmlog.find_one({"sn":sn,"isFail":False})
        return not result is None

    def isSnExistII(self,sn):
        result = self.cmlog.find_one({"sn":sn})
        return not result is None

    def insertCm(self,sn,mac):
        self.cmlog.update_one({"sn":sn},{'$set':{"sn":sn,"mac":mac,"isFail":"None","logMsg":"","testLogMsg":""}},upsert=True)

    def updataLog(self,sn,log,isFail,errorStep):
        self.cmlog.update_one({"sn":sn},{'$set':{"logMsg" :log,"isFail" : isFail,"errorStep" : errorStep}})

    def updataTestLog(self,sn,log,isFail,errorStep):
        self.cmlog.update_one({"sn":sn},{'$set':{"testLogMsg" :log,"isFail" : isFail,"errorStep" : errorStep}})

    def getConfig(self, section, key):
        config = ConfigParser.ConfigParser()
        config.read('config.conf')
        return config.get(section, key)