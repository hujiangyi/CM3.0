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
labelWidth = 8
entryWidth = 30
stepCount = 20
userName = 'topvision'
password = 'topvision'
class cm30 :
    def getConfig(self,section,key):
        config = ConfigParser.ConfigParser()
        config.read('config.conf')
        return config.get(section, key)
    def __init__(self):
        self.step = 1
        logPath = './log/' + datetime.datetime.now().strftime('%Y%m%d%H%M%S') + '/'
        os.makedirs(logPath)
        self.initLog(logPath)

        self.targetVersion = self.getConfig('image','targetVersion')
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
        self.isCheckRepeate = IntVar()
        self.isCheckRepeate.set(1)
        self.version.set('{}'.format(self.targetVersion))
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
        buttonF = Frame(midLeftF)
        self.connectCmB = Button(buttonF, text='连接CM',command=self.connectCm,state='normal')
        self.connectCmB.pack(side=LEFT)
        self.startB = Button(buttonF, text='开始测试',command=self.testStart,state='disabled')
        self.startB.pack(side=LEFT)
        self.ledTestSuccessB = Button(buttonF, text='LED全部点亮',command=self.ledSuccess,state='disabled')
        self.ledTestSuccessB.pack(side=LEFT)
        self.ledTestFailB = Button(buttonF, text='LED没有点亮',command=self.ledFail,state='disabled')
        self.ledTestFailB.pack(side=LEFT)
        self.resetTestB = Button(buttonF, text='Reset测试',command=self.resetTest,state='disabled')
        self.resetTestB.pack(side=LEFT)
        self.resetTestFailB = Button(buttonF, text='人工确认失败',command=self.resetTestFail,state='disabled')
        self.resetTestFailB.pack(side=LEFT)
        buttonF.pack(side=BOTTOM,anchor=E,expand=YES,fill=Y)
        midLeftF.pack(side=LEFT,anchor=NW,expand=YES,fill=Y)
        midRightF = Frame(midF)
        Label(midRightF, text="日志:",width=labelWidth,justify=LEFT,anchor=W).pack(side=TOP,anchor=W)
        self.logText = Text(midRightF,width=60,state=DISABLED)
        self.logText.pack(side=TOP,anchor=W,expand=YES,fill=BOTH)
        midRightF.pack(side=RIGHT,expand=YES,fill=Y)
        midF.pack(side=TOP,anchor=NW,expand=YES,fill=Y)
        frameRoot.pack(side=TOP,anchor=NW,expand=YES,fill=Y)
        root.mainloop()
    def connectCm(self):
        t = threading.Thread(target=self.connectCmThread)
        t.start()
    def connectCmThread(self):
        cmIp = self.getConfig('telnet','cmip')
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
        self.initData()
        t = threading.Thread(target=self.testStartThread)
        t.start()
    def testStartThread(self):
        cmIp = self.getConfig('telnet','cmip')
        port = self.getConfig('telnet','port')
        self.userName = self.getConfig('telnet','username')
        self.password = self.getConfig('telnet','password')
        isContinue = False
        try:
            #开始测试将按钮置灰
            self.startB.configure(state='disabled')
            self.telnet = Telnet(cmIp,port)
            self.readuntil('>',10)
            sn = self.readSn()
            if sn == None:
                return
            self.sn.set(sn)
            if not self.checkRepeated():
                return
            if not self.lockFreq():
                return
            if not self.checkOnline():
                return
            if not self.pingCC():
                return
            if not self.testLed():
                return
            isContinue = True
        except BaseException:
            print 'traceback.format_exc():\n%s' % traceback.format_exc()
            self.log('程序执行异常，测试终止',isFail=True)
        finally:
            if not isContinue:
                self.finishTest()
    def testContinue(self):
        isFail = not self.restoreFactory()
        self.finishTest(isFail=isFail)
    def readSn(self):
        self.send("docsDevSerialNumber")
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        lines = re.split('\r\n')
        for line in lines :
            if 'docsDevSerialNumber' in line:
                index = line.find(':') + 1
                sn = line[index:len(line)].strip()
                self.log('读取SN成功{}，测试继续'.format(sn))
                return sn
        self.log('读取SN不成功，测试终止',isFail=True)
        return None
    def checkRepeated(self):
        self.step = 1
        if self.db.isSnExistII(self.sn.get()) :
            self.log('CM SN存在，可以开始测试',isFail=True)
            return True
        else :
            self.log('CM SN不存在，测试终止')
            return False
    def lockFreq(self):
        self.step = 14
        lockFreq = self.getConfig('config','lockFreq')
        lockTryCount = self.getConfig('config','lockTryCount')
        freqs = lockFreq.split(',')
        for freq in freqs:
            isFail = True
            for i in range(1,int(lockTryCount) + 1):
                self.send('lock_ds {}'.format(freq))
                re = self.readuntil('Console/NonVol/Factory NonVol>')
                if 'QAM: lock, FEC: lock' in re :
                    isFail = False
                    self.log('锁频测试{}成功，测试继续'.format(freq))
                    break
                else :
                    self.log('第{}次锁频测试{}失败'.format(i,freq))
            if isFail :
                self.log('锁频测试{}重试超过次数，测试终止'.format(freq), isFail=True)
                return False
        return True


    def checkOnline(self):
        self.step = 15
        interval = self.getConfig('online_check','interval')
        times = self.getConfig('online_check','times')
        gotods = self.getConfig('online_check','gotods')
        self.send('goto_ds {}'.format(gotods))
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.sleepT(8)
        checkOnlineStatus = False
        for i in range(1,int(times)):
            self.sleepT(int(interval))
            self.send('show_cmStatus')
            re = self.readuntil('Console/NonVol/Factory NonVol>')
            if 'Docsis Registration Status: Operational' in re :
                self.log('CM上线成功，测试继续')
                checkOnlineStatus = True
                break
            else :
                if i % 10 == 0 :
                    self.log('CM还未成功上线')
        if not checkOnlineStatus:
            self.log('CM上线不成功，超过最大等待次数，测试终止',isFail=True)
            return False
        else :
            self.send('check_dsLocked')
            re = self.readuntil('Console/NonVol/Factory NonVol>')
            if 'Ds lock stauts: complete!' in re :
                self.log('CM成功绑定所有下行信道，测试继续')
            else :
                self.log('CM没有成功绑定所有下行信道，测试终止',isFail=True)
                return False
            self.send('check_usLocked')
            re = self.readuntil('Console/NonVol/Factory NonVol>')
            if 'Us lock stauts: complete!' in re :
                self.log('CM成功绑定所有上行信道，测试继续')
                return True
            else :
                self.log('CM没有成功绑定所有上行信道，测试终止',isFail=True)
                return False
    def pingCC(self):
        self.step = 16
        ccip = self.getConfig('ping','ccip')
        count = self.getConfig('ping','count')
        re = ping(ccip,count=int(count),udp=False)
        if re.packet_lost == 0:
            self.log('持续ping cc无丢包，测试继续')
            return True
        else:
            self.log('持续ping cc有丢包，测试终止{}'.format(re),isFail=True)
            return False

    def testLed(self):
        self.step = 17
        self.send('signal_test_on')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Signal test is enabled.' in re:
            self.log('开启信号测试模式成功，测试继续')
        else :
            self.log('开启信号测试模式失败，测试终止',isFail=True)
            return False

        self.send('led all on')
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.log('已经打开所有LED，请观察，并选择测试结果')
        self.ledTestSuccessB.config(state='normal')
        self.ledTestFailB.config(state='normal')
        return True
    def ledSuccess(self):
        self.log('开启信号测试模式成功，测试继续')
        self.ledTestSuccessB.config(state='disabled')
        self.ledTestFailB.config(state='disabled')
        self.send('led all off')
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.resetTestB.config(state='normal')
    def ledFail(self):
        self.log('LED灯未全部打开，测试终止',isFail=True)
        self.ledTestSuccessB.config(state='disabled')
        self.ledTestFailB.config(state='disabled')
        self.send('led all off')
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.send('signal_test_off')
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.finishTest()
    def resetTest(self):
        self.log('Reset按键测试开始，请在10s内按下CM的复位键')
        t = threading.Thread(target=self.resetTestThread)
        t.start()
    def resetTestThread(self):
        self.step = 18
        self.resetTestB.config(state='disabled')
        self.resetTestFailB.config(state='disabled')
        self.send('button_test')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Reset button test OK.' in re:
            self.log('Reset按键测试成功，测试继续')
            self.send('signal_test_off')
            re = self.readuntil('Console/NonVol/Factory NonVol>')
            self.testContinue()
        else :
            self.log('Reset按键测试失败，重新测试或人工确认测试失败')
            self.resetTestB.config(state='normal')
            self.resetTestFailB.config(state='normal')
    def resetTestFail(self):
        self.log('人工确认Reset按键失效，测试终止',isFail=True)
        self.resetTestB.config(state='disabled')
        self.resetTestFailB.config(state='disabled')
        self.send('signal_test_off')
        self.readuntil('Console/NonVol/Factory NonVol>')
        self.finishTest()
    def restoreFactory(self):
        self.step = 19
        self.send('restore_defaults')
        re = self.readuntil('Console/NonVol/Factory NonVol>')
        if 'Set Ok!' in re:
            self.log('恢复出厂设置成功，测试结束')
            return True
        else:
            self.log('恢复出厂设置失败，测试终止')
            return False

    def finishTest(self,isFail = True):
        self.logText.configure(state=NORMAL)
        testLog = self.logText.get(1.0,END)
        self.logText.configure(state=DISABLED)
        if isFail :
            self.db.updataTestLog(self.sn.get(),testLog,isFail ,"{}".format(self.step))
        else :
            self.db.updataTestLog(self.sn.get(),testLog,isFail,"{}".format(20))
        self.step = 20
        self.log('测试结束',isFail=isFail)
        self.close()
        self.saveLog()
        self.connectCmB.configure(state='normal')
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
        self.ledTestSuccessB.configure(state='disabled')
        self.ledTestFailB.configure(state='disabled')
        self.resetTestB.configure(state='disabled')
        self.resetTestFailB.configure(state='disabled')
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
cm30()
