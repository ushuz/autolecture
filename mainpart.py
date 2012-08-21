# -*- coding:utf-8 -*-
import time

import win32com.client
from const import Const

class Class:
    u'''Class类 代表一门课程
    包含了一门课程location、subject、duration、startTime等属性以及save等方法'''
    def __init__(self,inString):
        u''''''
        self.rawList = inString.split(' ')
        self.dataDict = {}
        self.__parse()
        self.__getExtraWeeks()
    
    def __parse(self):
        u'''Class.__parse私有方法
    从已拆分好的rawList中解析出location、subject、duration、startTime、weekDayNo等属性
    并储存到dataDict'''
        self.dataDict['subject'] = self.rawList[1]
        self.dataDict['location'] = Const.ClassRoomBuilding[int(self.rawList[2][0])] + self.rawList[2][1:]
        self.dataDict['weekDayNo'] = self.rawList[0][0]
        if u'二教' in self.dataDict['location'] and self.rawList[0][1:] == '34':
            self.dataDict['duration'] = 100
            self.dataDict['startTime'] = '10:05'
        else:
            self.dataDict['duration'] = 110
            self.dataDict['startTime'] = Const.ClassBeginTime[self.rawList[0][1:]]
        
        if debug:
            print u"\n打印['location'] ['duration'] ['startTime'] ['subject']"#FORDEBUG
            print self.dataDict['location'],self.dataDict['duration'],self.dataDict['startTime'],self.dataDict['subject']#FORDEBUG

    def __getExtraWeeks(self):
        u'''Class.__getRecurrWeeks私有方法
    获取无课周次以便之后删除'''
        tmp = self.rawList[3].split(',')

        #将 'A-B' 形式的recurrence weeks展开
        for i in tmp:
            if '-' in i:
                t = i.split('-')
                tmp += [str(j) for j in range(int(t[0]), int(t[1])+1)]         #使用list comprehension将list中每个元素都转变为str
                tmp[tmp.index(i)] = None                                  #展开完成后清空原位置元素

        tmp = [int(i) for i in tmp if i is not None]                        #将tmp里的元素都转换为int并去掉值为None的元素以方便排序
        tmp.sort()

        allWeeks = [str(i) for i in range(tmp[0], tmp[-1]+1)]               #allWeeks的元素均为str方便后续操作

        self.dataDict['extraWeeks'] = [i for i in allWeeks if int(i) not in tmp]          #无课的周次
        self.recurrWeeks = [str(i) for i in tmp]                                        #上课的周次
        
        self.dataDict['startDate'] = C.DateDict[int(allWeeks[0] + self.rawList[0][0])]      #第一次上课日期    即recurrencePatternStartDate
        self.dataDict['endDate'] = C.DateDict[int(allWeeks[-1] + self.rawList[0][0])]       #最后一次上课日期  即recurrencePatternEndDate
        
        
        
        if debug:
            for k,v in self.dataDict.items():
                print k,v
            print u"\n打印['extraWeeks']"#FORDEBUG
            print self.dataDict['extraWeeks']#FORDEBUG
    
    def save(self):
        u'''save方法
    生成课程的appointment并储存到Outlook'''            
        # TODO
        try:
            apptGen(self.dataDict)
            return True
        except Exception, e:
            if debug:
                print e
            return


def apptGen(dataDict):
    u'''apptGen()函数
    从传入的dataDict参数中获取具体属性以生成指定的AppointmentItem'''
    o = win32com.client.Dispatch("Outlook.Application")     #新建一个Outlook.Application实例 o
    a = o.CreateItem(1)                                     #新建一个AppointmentItem实例 a

    a.Start = dataDict['startDate']+' '+dataDict['startTime']    #设置课程开始时间
    a.Duration = dataDict['duration']                       #设置课程时长
    a.Subject = dataDict['subject']                         #设置课程名称
    a.Location = dataDict['location']                       #设置授课教室
    a.ReminderSet = False

    p = a.GetRecurrencePattern()                            #获取AppointmentItem.GetRecurrencePattern()对象以修改a的重复类型
    p.RecurrenceType = 1                                    #设置重复类型为 1 == olRecursWeekly

    p.PatternStartDate = dataDict['startDate']              #设置重复开始日期
    p.PatternEndDate = dataDict['endDate']                  #设置重复结束日期

    a.Save()                                                #储存 a
    
    #之前创建了约会系列，之后删除系列中的无效约会
    i = o.Session.GetDefaultFolder(9).Items
    s = i.Find("[Subject]='%s'" % dataDict['subject'])
    
    #确保s是最后该课程最新系列
    #防止名称相同的旧课程系列乱入
    while True:
        t = i.FindNext()
        if t is not None:
            s = t
            if debug:
                print s.Start
            continue
        break;
            
    sp = s.GetRecurrencePattern()
    
    for i in dataDict['extraWeeks']:
        tofind = C.DateDict[int(i + dataDict['weekDayNo'])]+' '+dataDict['startTime']
        if debug:
            print u"\n将删除", tofind, u"的", dataDict['subject']
        todelete = sp.GetOccurrence(time.strptime(tofind, "%Y-%m-%d %H:%M"))
        todelete.Delete()
        


def init():
    '''init函数
    初始化部分'''
    show('welcome')                                                 #显示欢迎信息

    if not debug:
        show(0)                                                         #获取
        termFirstDayStr = raw_input()                                   #学期第一周周一日期 --> string
        show(1)                                                         #获取
        termTotalWeeks = int(raw_input())                               #学期总周数 --> integer
    else:
        termFirstDayStr = "20120903"
        termTotalWeeks = 20

    C.genDateDict(termFirstDayStr,termTotalWeeks)                   #调用C的genDateDict方法生成DateDict


def show(s):
    '''show函数
    打印指定的提示信息'''
    if s == 'welcome':
        print u'''AutoLecture
欢迎使用
程序将帮助你根据课程表快速生成一系列Outlook appointments'''
        print '\n'
    elif s == 0:
        print u"请输入学期第一周周一的日期(str.'YYYYMMDD')"
    elif s == 1:
        print u'请输入学期总周数(int.dd)'
    elif s == 2:
        print u'''请输入课程信息(str.'WSS SUBJECT LOCATION RECURRWEEKS')
eg. '112 高等数学 1108 2-11,14,17'"
    周一1-2节 高等数学 一教108 2-11,14,17周上课'''
    elif s == 'debug':
        print '###FORDEBUG###'


if __name__ == '__main__':
    debug = False
    C = Const()                                                                 #创建Const实例C以提供相应常量
    
    init()
    
    if debug:
        show('debug')#FORDEBUG
        print C.DateDict                                                    #打印weekDayDict
        show('debug')#FORDEBUG

    show(2)

    while True:
        tmp = raw_input()
        try:
            tmp = tmp.decode("cp936")
        except UnicodeDecodeError:
            tmp = tmp.decode("utf-8")
        except UnicodeDecodeError:
            tmp = tmp.decode("gbk")
        except UnicodeDecodeError:
            tmp = tmp.decode("gb2312")
        except UnicodeDecodeError:
            tmp = tmp.decode("ascii")
        except:
            print u"未知文字编码，程序退出。"
            break
        print type(tmp)
        if tmp == u'':
            break
        else:
            one = Class(tmp)
            if one.save():
                print u"保存成功"
