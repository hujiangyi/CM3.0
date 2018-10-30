#! /usr/bin/env python
# -*- coding:utf-8 -*-
# vim:fenc=utf-8
#
# Copyright © 2014 jay <hujiangyi@dvt.dvt.com>
#

from Tkinter import *
import ConfigParser
import xlrd
from tkMessageBox import *
from telnetlib import *
from pyping import *
import traceback
import datetime
import threading
import sys
from database import *

reload(sys)
sys.setdefaultencoding('gbk')

# 设置窗口大小
width = 800
height = 400
labelWidth = 12
entryWidth = 30
stepCount = 20
userName = 'topvision'
password = 'topvision'
class cm30forset :
    def getConfig(self,section,key):
        config = ConfigParser.ConfigParser()
        config.read('config.conf')
        return config.get(section, key)
    def fixMac(self,event):
        sn = self.sn.get()
        rb = xlrd.open_workbook('sntomac.xlsx')
        sheetCount = len(rb.sheets())
        for si in range(sheetCount):
            sheetR = rb.sheet_by_index(si)
            nrows = sheetR.nrows
            for row in range(1,nrows):
                key = self.getStringValue(sheetR,row, 0)
                mac1 = self.getStringValue(sheetR,row, 1)
                if sn == key :
                    self.mac1.set(mac1)
                    self.initData()
                    self.connectCmB.configure(state='normal')
                    return
        showerror('错误', '没有找到对应的MAC，请检查配置文件')
        self.mac1.set('')
        self.initData()
        self.connectCmB.configure(state=DISABLED)
    def __init__(self):
        self.step = 1
        logPath = './log/' + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '/'
        os.makedirs(logPath)
        self.initLog(logPath)

        self.targetVersion = self.getConfig('image','targetVersion')
        self.targetProductModel = self.getConfig('image','targetProductModel')
        self.db = database()
        root = Tk()
        # 设置标题
        root.title('Python GUI Learning')
        # 获取屏幕尺寸以计算布局参数，使窗口居屏幕中央
        screenwidth = root.winfo_screenwidth()
        screenheight = root.winfo_screenheight()
        alignstr = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
        root.geometry(alignstr)
        # 设置窗口是否可变长、宽，True：可变，False：不可变
        root.resizable(width=False, height=True)

        self.sn = StringVar()
        self.mac1 = StringVar()
        self.version = StringVar()
        self.productModel = StringVar()
        self.isCheckRepeate = IntVar()
        self.isCheckRepeate.set(1)
        self.version.set('{}'.format(self.targetVersion))
        self.productModel.set('{}'.format(self.targetProductModel))
        self.stateLabels = []
        frameRoot = Frame(root)
        topF = Frame(frameRoot)
        for i in range(1,stepCount + 1):
            label = Label(topF, text="{}".format(i))
            label.pack(side=LEFT)
            self.stateLabels.append(label)

        topF.pack(side=TOP,anchor=NW)
        midF = Frame(frameRoot)
        midLeftF = Frame(midF)
        Label(midLeftF, text="参数:",width=labelWidth,justify=LEFT,anchor=W).pack(side=TOP,anchor=W)
        snF = Frame(midLeftF)
        Label(snF, text="SN:",width=labelWidth,justify=LEFT,anchor=W).pack(side=LEFT)
        snEntry = Entry(snF,width=entryWidth, textvariable=self.sn)
        snEntry.pack(side=LEFT)
        snF.pack(side=TOP,anchor=W)
        mac1F = Frame(midLeftF)
        Label(mac1F, text="MAC1:",width=labelWidth,justify=LEFT,anchor=W).pack(side=LEFT)
        Entry(mac1F,width=entryWidth, textvariable=self.mac1,state='readonly').pack(side=LEFT)
        mac1F.pack(side=TOP,anchor=W)
        versionF = Frame(midLeftF)
        Label(versionF, text="Version:",width=labelWidth,justify=LEFT,anchor=W).pack(side=LEFT)
        Entry(versionF,width=entryWidth, textvariable=self.version,state='readonly').pack(side=LEFT)
        versionF.pack(side=TOP,anchor=W)
        productModelF = Frame(midLeftF)
        Label(productModelF, text="Product Model:",width=labelWidth,justify=LEFT,anchor=W).pack(side=LEFT)
        Entry(productModelF,width=entryWidth, textvariable=self.productModel,state='readonly').pack(side=LEFT)
        productModelF.pack(side=TOP,anchor=W)
        checkBoxF = Frame(midLeftF)
        Checkbutton(checkBoxF, text="是否检查SN和MAC重复", variable=self.isCheckRepeate).pack(side=LEFT)
        checkBoxF.pack(side=TOP,anchor=W)
        buttonF = Frame(midLeftF)
        self.connectCmB = Button(buttonF, text='连接CM',command=self.connectCm,state='disabled')
        self.connectCmB.pack(side=LEFT)
        self.startB = Button(buttonF, text='开始设置',command=self.testStart,state='disabled')
        self.startB.pack(side=LEFT)
        buttonF.pack(side=BOTTOM,anchor=E,expand=YES,fill=Y)
        midLeftF.pack(side=LEFT,anchor=NW,expand=YES,fill=Y)
        midRightF = Frame(midF)
        Label(midRightF, text="日志:",width=labelWidth,justify=LEFT,anchor=W).pack(side=TOP,anchor=W)
        self.logText = Text(midRightF,width=60,state=DISABLED)
        self.logText.pack(side=TOP,anchor=W,expand=YES,fill=BOTH)
        midRightF.pack(side=RIGHT,expand=YES,fill=Y)
        midF.pack(side=TOP,anchor=NW,expand=YES,fill=Y)
        frameRoot.pack(side=TOP,anchor=NW,expand=YES,fill=Y)
        snEntry.bind('<Key-Return>', self.fixMac)
        root.mainloop()
    def checkRepeated(self):
        self.step = 1
        if self.isCheckRepeate.get() ==1 :
            if self.db.isMacExist(self.mac1.get()) :
                self.log('CM MAC已经使用，不能开始测试',isFail=True)
                return False
            if self.db.isSnExist(self.sn.get()) :
                self.log('CM SN已经使用，不能开始测试',isFail=True)
                return False
        else :
            self.log('跳过MAC和SN重复检查，继续测试')
        self.log('MAC和SN重复检查通过，继续测试')
        self.db.insertCm(self.sn.get(),self.mac1.get())
        return True
    def connectCm(self):
        t = threading.Thread(target=self.connectCmThread)
        t.start()
    def connectCmThread(self):
        cmIp = self.getConfig('telnet','cmip')
        if not self.checkRepeated():
            return
        self.step = 2
        re = ping(cmIp,udp=False)
        if re.ret_code == 0 :
            self.log('CM能ping通，可以开始测试')
            self.startB.configure(state='normal')
            self.connectCmB.configure(state='disabled')
        else:
            self.log('CM不能ping通，不能开始测试',isFail=True)
            self.startB.configure(state='disabled')
            self.finishTest()
    def testStart(self):
        t = threading.Thread(target=self.testStartThread)
        t.start()
    def testStartThread(self):
        cmIp = self.getConfig('telnet','cmip')
        port = self.getConfig('telnet','port')
        self.userName = self.getConfig('telnet','username')
        self.password = self.getConfig('telnet','password')
        try:
            #开始测试将按钮置灰
            self.startB.configure(state='disabled')
            self.telnet = Telnet(cmIp,port)
            self.readuntil('>',10)
            if not self.checkVersion():
                return
            if not self.checkProductModel():
                return
            if not self.writeMAC():
                return
            if not self.writeSN():
                return
            if not self.writeFreq():
                return
            if not self.writeAnnex():
                return
            if not self.writeIPSV():
                return
            if not self.checkSn():
                return
            if not self.checkMac():
                return
            if not self.checkAnnex():
                return
            if not self.checkIPSV():
                return
            if not self.checkFreq():
                return
            self.finishTest(isFail=False)
        except BaseException:
            print 'traceback.format_exc():\n%s' % traceback.format_exc()
            self.log('程序执行异常，测试终止',isFail=True)
            self.finishTest()
    def checkVersion(self):
        self.step = 3
        self.send('version')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        res = re.split('\r\n')
        for s in res:
            if 'Software Version' not in s:
                continue
            vs = s.split(':')
            version = vs[1].strip()
            if self.targetVersion != version:
                self.log('当前CM不是匹配版本，测试终止',isFail=True)
                return False
            else :
                self.log('当前CM是匹配版本，测试继续')
                return True
        self.log('没有读取到当前CM的软件版本，测试终止',isFail=True)
        return False
    def checkProductModel(self):
        self.step = 3
        self.send('version')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        res = re.split('\r\n')
        for s in res:
            if 'Product Model' not in s:
                continue
            vs = s.split(':')
            productModel = vs[1].strip()
            if self.targetProductModel != productModel:
                self.log('当前CM不是匹配的产品型号，测试终止',isFail=True)
                return False
            else :
                self.log('当前CM是匹配的产品型号，测试继续')
                return True
        self.log('没有读取到当前CM的的产品型号，测试终止',isFail=True)
        return False
    def writeMAC(self):
        self.step = 4
        self.send('mac_address 1 {}'.format(self.mac1.get()))
        re = self.readuntil('Console/NonVol/Factory NonVol>')

        if 'Set Ok!' in re :
            self.log('烧写MAC1成功[{}]，测试继续'.format(self.mac1.get()))
            return True
        else :
            self.log('烧写MAC1不成功[{}]，测试终止'.format(self.mac1.get()),isFail=True)
            return False
    def checkMac(self):
        self.step = 10
        if self.check('mac_address 1','MAC address for IP stack 1',self.mac1.get()):
            self.log('检查MAC正确，继续测试')
            return True
        else :
            self.log('检查MAC错误，终止测试',isFail=True)
            return False
    def writeSN(self):
        self.step = 5
        self.send('docsDevSerialNumber {}'.format(self.sn.get()))
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Set Ok!' in re :
            self.log('烧写SN成功，测试继续')
            return True
        else :
            self.log('烧写SN不成功，测试终止',isFail=True)
            return False
    def checkSn(self):
        self.step = 9
        if self.check('docsDevSerialNumber','docsDevSerialNumber',self.sn.get()):
            self.log('检查SN正确，继续测试')
            return True
        else :
            self.log('检查SN错误，终止测试',isFail=True)
            return False

    def writeFreq(self):
        self.step = 6
        self.send('ds_freq 0')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Set Ok!' in re :
            self.log('清空预配置频点成功，测试继续')
        else :
            self.log('清空预配置频点失败，测试终止',isFail=True)
            return False
        preFreq = self.getConfig('config','preFreq')
        freqs = preFreq.split(',')
        for freq in freqs:
            self.send('ds_freq {}'.format(freq))
            re = self.readuntil('Console/NonVol/Factory NonVol>')
            if 'Set Ok!' in re:
                self.log('写入预配置频点{}成功，测试继续'.format(freq))
            else:
                self.log('写入预配置频点{}失败，测试终止'.format(freq),isFail=True)
                return False
        self.log('完成预配置频点写入')
        return True

    def checkFreq(self):
        self.step = 13
        preFreq = self.getConfig('config','preFreq')
        freqs = preFreq.split(',')
        fs = []
        for freq in freqs:
            fs.append(freq.strip())
        cmdFs = []
        self.send('ds_freq')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        lines = re.split('\r\n')
        for line in lines:
            if '('not in line and  ')' in line :
                cols = line.split(')')
                cmdFs.append(cols[1].strip())
        for freq in fs:
            if freq not in cmdFs:
                self.log('检查频点{}失败，未配置，终止测试'.format(freq),isFail=True)
                return False
        for freq in cmdFs:
            if freq not in fs:
                self.log('检查频点{}失败，未清除，终止测试'.format(freq),isFail=True)
                return False
        self.log('检查频点正常，测试继续')
        return True
    def writeAnnex(self):
        self.step = 7
        cm_annex = self.getConfig('config','cm_annex')
        if cm_annex == '0':
            annexStr = '美标'
        elif cm_annex == '1':
            annexStr = '欧标'
        else:
            self.log('错误的欧美标配置[{}]，测试终止'.format(cm_annex),isFail=True)
            return False
        self.send('cm_annex {}'.format(cm_annex))
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Set Ok!' in re :
            self.log('欧美标配置[{}]成功，测试继续'.format(annexStr))
            return True
        else :
            self.log('欧美标配置[{}]不成功，测试终止'.format(annexStr),isFail=True)
            return False
    def checkAnnex(self):
        self.step = 11
        cm_annex = self.getConfig('config','cm_annex')
        if self.check('cm_annex','CM DOCSIS Annex Mode',cm_annex):
            self.log('检查欧美标模式正确，继续测试')
            return True
        else :
            self.log('检查欧美标模式错误，终止测试',isFail=True)
            return False
    def writeIPSV(self):
        self.step = 8
        ipsvNum = self.getConfig('config','ipsvNum')
        self.send('ipsv_chan_num {}'.format(ipsvNum))
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Set Ok!' in re :
            self.log('IPSV信道配置[{}]成功，测试继续'.format(ipsvNum))
            return True
        else :
            self.log('IPSV信道配置[{}]不成功，测试终止'.format(ipsvNum),isFail=True)
            return False
    def checkIPSV(self):
        self.step = 12
        ipsvNum = self.getConfig('config','ipsvNum')
        if self.check('ipsv_chan_num','IPSV channel number',ipsvNum):
            self.log('检查IPSV信道个数配置正确，继续测试')
            return True
        else :
            self.log('检查IPSV信道个数配置错误，终止测试',isFail=True)
            return False

    def finishTest(self,isFail = True):
        self.logText.configure(state=NORMAL)
        testLog = self.logText.get(1.0,END)
        self.logText.configure(state=DISABLED)
        if isFail :
            self.db.updataLog(self.sn.get(),testLog,isFail ,"{}".format(self.step))
        else :
            self.db.updataLog(self.sn.get(),testLog,isFail,"{}".format(20))
        self.step = 20
        self.log('测试结束',isFail=isFail)
        self.close()
        self.saveLog()
    ###############################################com##################################################################
    def initLog(self, logPath):
        self.logPath = logPath
        self.cmdResultFile = open(logPath  + "CmdResult.log", "w")
        self.logResultFile = open(logPath  + "logFile.log", "w")
    def initData(self):
        for i in range(1,stepCount + 1):
            self.stateLabels[i -1].configure(fg='black')
        self.connectCmB.configure(state='disabled')
        self.startB.configure(state='disabled')
        self.logText.configure(state=NORMAL)
        self.logText.delete(1.0,END)
        self.logText.configure(state=DISABLED)

    def close(self):
        try:
            self.telnet.close()
        except Exception, msg:
            pass

    def send(self, cmd):
        terminator = '\r'
        cmd = str(cmd)
        cmd += terminator
        try:
            msg = cmd
            self.telnet.write(cmd)
            self.cmdLog(cmd)
        except Exception, msg:
            self.log(`msg`)
            raise Exception("telnet write error!")

    def sendII(self, cmd):
        cmd = str(cmd)
        try:
            msg = cmd
            self.telnet.write(cmd)
        except Exception, msg:
            self.log(`msg`)
            raise Exception("telnet write error!")

    def read(self):
        str = self.telnet.read_very_eager()
        self.cmdLog(str)
        return str

    def readuntil(self, waitstr='xxx', timeout=0):
        tmp = ""
        if timeout != 0:
            delay = 0.0
            while delay <= timeout:
                tmp += self.read()
                if self.needLogin(tmp):
                    tmp = ''
                    self.send('')
                    tmp += self.read()
                time.sleep(1)
                if tmp.endswith('--More--'):
                    self.sendII(' ')
                if waitstr in tmp:
                    return tmp
                delay += 1
            raise Exception("wait str timeout")
        else:
            while True:
                tmp += self.read()
                # self.log(tmp)
                if self.needLogin(tmp):
                    tmp = ''
                    self.send('')
                    tmp += self.read()
                if waitstr in tmp:
                    return tmp

    def readuntilMutl(self, waitstrs=['xxx'], timeout=0):
        tmp = ""
        if timeout != 0:
            delay = 0.0
            while delay <= timeout:
                time.sleep(1)
                tmp += self.read()
                if tmp.endswith('--More--'):
                    self.sendII(' ')
                for waitstr in waitstrs:
                    if waitstr in tmp:
                        return tmp
                delay += 1
            raise Exception("wait str timeout")
        else:
            while True:
                tmp += self.read()
                # self.log(tmp)
                if self.needLogin(tmp):
                    tmp = ''
                    self.send('')
                    tmp += self.read()
                for waitstr in waitstrs:
                    if waitstr in tmp:
                        return tmp

    def readuntilII(self, waitstr='xxx', timeout=0):
        tmp = ""
        if timeout != 0:
            delay = 0.0
            while delay <= timeout:
                time.sleep(1)
                tmp += self.read()
                if waitstr in tmp:
                    return tmp
                delay += 1
            raise Exception("wait str timeout")
        else:
            while True:
                tmp += self.read()
                if waitstr in tmp:
                    return tmp

    def needLogin(self, str):
        try:
            if 'Login:' in str:
                self.send(self.userName)
                self.readuntilII(waitstr='Password:', timeout=30)
                self.send(self.password)
                self.readuntilII('>', timeout=30)
                self.send('cd /non-vol/factory')
                self.readuntilII('>', timeout=30)
                return True
            else:
                return False
        except Exception, msg:
            self.log(`msg`)
            raise Exception("login faild!")

    def sleepT(self, delay):
        time.sleep(delay)

    def cmdLog(self, str):
        self.cmdResultFile.write(str)
        self.cmdResultFile.flush()

    def log(self, str,isFail=False):
        print str
        str = '{}\t{}\r\n'.format(datetime.datetime.now().strftime('%Y%m%d%H%M%S'),str)
        self.logResultFile.write(str)
        self.logResultFile.flush()
        self.logText.configure(state=NORMAL)
        self.logText.insert(END,str)
        self.logText.configure(state=DISABLED)
        if not isFail:
            self.stateLabels[self.step -1].configure(fg='green')
        else :
            self.stateLabels[self.step-1].configure(fg='red')

    def getStringValue(self, sheet,row, col):
        ctype = sheet.cell(row, col).ctype
        cell = sheet.cell_value(row, col)

        if ctype == 2 and cell % 1 == 0.0:
            cell = int(cell)
        return '{}'.format(cell)
    def check(self,cmd,key,value):
        self.send(cmd)
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        lines = re.split('\r\n')
        for line in lines :
            if key in line:
                index = line.find(':') + 1
                col = line[index:len(line)].strip()
                return col == value
        return False
    def saveLog(self):
        pass
cm30forset()
