# -*- coding:utf-8 -*-
import time
import datetime
import logging

import win32com.client


BUILDINGS = {"1": u'一教', "2": u'二教', "3": u'三教', "4": u'四教'}                                  #授课教学楼
START_TIMES = {'1':'8:00', '2':'10:10', '3':'13:30', '4':'15:40', '5':'19:00'}                        #课程开始时间
DATES = {}


class Class:
    u'''Class类 代表一门课程
    包含了一门课程location、subject、duration、startTime等属性以及save等方法'''
    def __init__(self,inString):
        u''''''
        self.rawList = inString.split(' ')
        self.dataDict = {}
        self._parse()
        self._getExtraWeeks()
        # self.save()
    
    def _parse(self):
        u'''Class._parse私有方法
    从已拆分好的rawList中解析出location、subject、duration、startTime、weekDayNo等属性
    并储存到dataDict'''
        self.dataDict['subject'] = self.rawList[1]
        if self.rawList[3][0] in BUILDINGS.keys():
            self.dataDict['location'] = BUILDINGS[self.rawList[3][0]] + self.rawList[3][1:]
        else:
            self.dataDict["location"] = self.rawList[3]
        self.dataDict['weekDayNo'] = self.rawList[0][0]
        if u'二教' in self.dataDict['location'] and self.rawList[0][1:] == '34':
            self.dataDict['duration'] = 100
            self.dataDict['startTime'] = '10:05'
        else:
            self.dataDict['duration'] = 110
            self.dataDict['startTime'] = START_TIMES[self.rawList[0][1:]]
        
    def _getExtraWeeks(self):
        u'''Class._getExtraWeeks私有方法
    获取无课周次(extraWeeks)、系列开始日期(startDate)和结束日期(endDate)'''
        tmp = self.rawList[2].split(',')

        #将 'A-B' 形式的recurrence weeks展开
        for i in tmp:
            if '-' in i:
                t = i.split('-')
                tmp += [str(j) for j in xrange(int(t[0]), int(t[1])+1)]                 #使用list comprehension将list中每个元素都转变为str
                tmp[tmp.index(i)] = None                                                #展开完成后清空原位置元素

        tmp = [int(i) for i in tmp if i is not None]                                    #将tmp里的元素都转换为int并去掉值为None的元素以排序
        tmp.sort()

        allWeeks = [str(i) for i in xrange(tmp[0], tmp[-1]+1)]                          #allWeeks的元素均为str方便后续操作

        self.dataDict['extraWeeks'] = [i for i in allWeeks if int(i) not in tmp]        #无课的周次
        
        self.dataDict['startDate'] = DATES[int(allWeeks[0] + self.rawList[0][0])]       #第一次上课日期    即recurrencePatternStartDate
        self.dataDict['endDate'] = DATES[int(allWeeks[-1] + self.rawList[0][0])]        #最后一次上课日期  即recurrencePatternEndDate

        for k,v in self.dataDict.items():
            log(u"%s: %s" % (k, v))

    def save(self):
        u'''Class.save方法
    生成该课程的appointment并储存到Outlook'''
        try:
            apptGen(self.dataDict)
            print u"创建成功: %s" % self.dataDict["subject"]
        except Exception, e:
            logging.error(e)
            print u"创建失败: %s" % self.dataDict["subject"]
            return


def apptGen(dataDict):
    u'''apptGen()函数
    从传入的dataDict参数中获取具体属性以生成指定的AppointmentItem'''
    o = win32com.client.Dispatch("Outlook.Application")     #新建一个Outlook.Application实例 o
    a = o.CreateItem(1)                                     #新建一个AppointmentItem实例 a

    #补完实例 a 的各种属性
    a.Start = dataDict['startDate']+' '+dataDict['startTime']
    a.Duration = dataDict['duration']
    a.Subject = dataDict['subject']
    a.Location = dataDict['location']
    a.ReminderSet = False

    #创建约会系列
    #每周一次，从该课程第一周到最后一周
    p = a.GetRecurrencePattern()                            #获取AppointmentItem.GetRecurrencePattern()对象以修改a的重复类型
    p.RecurrenceType = 1                                    #设置重复类型为 1 == olRecursWeekly
    p.PatternStartDate = dataDict['startDate']
    p.PatternEndDate = dataDict['endDate']
    a.Save()
    log(u"已创建: %s %s\t多余课程: %d节" % (a.Start, a.Subject, len(dataDict["extraWeeks"])))
    
    #删除系列中的无效约会
    i = o.Session.GetDefaultFolder(9).Items
    s = i.Find("[Subject]='%s'" % dataDict['subject'])
    
    #确保s是最后该课程最新系列
    #防止名称相同的旧课程系列乱入
    while True:
        t = i.FindNext()
        if t is not None:
            s = t
            continue
        break;

    sp = s.GetRecurrencePattern()
    
    for i in dataDict['extraWeeks']:
        tofind = DATES[int(i + dataDict['weekDayNo'])]+' '+dataDict['startTime']
        todelete = sp.GetOccurrence(time.strptime(tofind, "%Y-%m-%d %H:%M"))
        todelete.Delete()
        log(u"删除多余课程: "+tofind+u" "+dataDict['subject'])
        
# def show(s):
    # '''show函数
    # 打印指定的提示信息'''
    # if s == 'welcome':
        # print u'''AutoLecture
# 欢迎使用
# 程序将帮助你根据课程表快速生成一系列Outlook appointments'''
        # print '\n'
    # elif s == 0:
        # print u"请输入学期第一周周一的日期(str.'YYYYMMDD')"
    # elif s == 1:
        # print u'请输入学期总周数(int.dd)'
    # elif s == 2:
        # print u'''请输入课程信息(str.'WSS SUBJECT RECURRWEEKS LOCATION')
# eg. '11 高等数学 1108 2-11,14,17'"
    # 周一1-2节 高等数学 2-11,14,17周 一教108上课'''
    # elif s == 'debug':
        # print '###FORDEBUG###'

def genDateDict(termFirstDayStr,termTotalWeeks):
    u'''genDateDict函数
由传入的termFirstDayStr和termTotalWeeksInt参数\n生成以(AB(A:digit,周目;B:digit,日目))为键的工作日词典常量DATES'''
    termFirstDay = datetime.date(int(termFirstDayStr[:4]),int(termFirstDayStr[4:6]),int(termFirstDayStr[6:]))

    weekDayList = []
    weekNo = 11             #表示第一周周一

    #生成工作日(weekday)日期列表(list)
    for i in xrange(0,termTotalWeeks*7):
        if not((i-5)%7 == 0 or (i-6)%7 == 0):
            weekDayList += [str(termFirstDay + datetime.timedelta(i))]

    #生成工作日(weekday)日期词典(dict)，以(AB(A:周目;B:日目))为键(key)
    global DATES
    for i in xrange(1,termTotalWeeks*5+1):
        DATES[weekNo] = weekDayList[i-1]
        if i % 5 == 0:
            weekNo += 6
        else:
            weekNo += 1

def main(termFirstDayStr, termTotalWeeks, lectures):
    #生成DATES常量
    
    genDateDict(termFirstDayStr, termTotalWeeks)
    
    for i in lectures:
        Class(i.decode("utf-8")).save()

if __name__ == '__main__':
    
    debug = True
    
    if debug:
        logging.basicConfig(level=logging.DEBUG,
                    format='%(levelname)s - %(message)s',
                    )
    else:
        logging.basicConfig(level=logging.ERROR,
                    format='%(levelname)s - %(message)s',
                    )
    log = logging.debug
    
    
    main("20120903", 20, [])