from const import Const

class Class:
    u'''Class类 代表一门课程
    包含了一门课程location、subject、duration、startTime等属性以及save等方法'''
    def __init__(self,inString):
        u''''''
        self.rawList = inString.split(' ')
        self.dataDict = {}
        self.__parse()
        self.__getRecurrWeeks()
        self.save()
    
    def __parse(self):
        u'''Class.__parse私有方法
    从已拆分好的rawList中解析出location、subject、duration、startTime等属性
    并储存到dataDict'''
        self.dataDict['subject'] = self.rawList[1]
        self.dataDict['location'] = Const.ClassRoomBuilding[int(self.rawList[2][0])] + self.rawList[2][1:]
        if u'二教' in self.dataDict['location'] and self.rawList[0][1:] == '34':
            self.dataDict['duration'] = 100
            self.dataDict['startTime'] = '10:05'
        else:
            self.dataDict['duration'] = 110
            self.dataDict['startTime'] = Const.ClassBeginTime[self.rawList[0][1:]]
        
        print u"\n打印['location'] ['duration'] ['startTime'] ['subject']"#FORDEBUG
        print self.dataDict['location'],self.dataDict['duration'],self.dataDict['startTime'],self.dataDict['subject']#FORDEBUG

    def __getRecurrWeeks(self):
        u'''Class.__getRecurrWeeks私有方法
    获取上课周次'''
        tmp = self.rawList[3].split(',')
        
        #将 'A-B' 形式的recurrence weeks展开
        for i in tmp:
            if '-' in i:
                t = i.split('-')
                tmp += [j for j in range(int(t[0]),int(t[1])+1)]   #使用list comprehension将list中每个元素都转变为str
                tmp[tmp.index(i)] = None                                #展开完成后清空原位置元素
        self.recurrWeeks = [int(i) for i in tmp if i is not None]
        self.recurrWeeks.sort()
        self.recurrWeeks = [str(i) for i in self.recurrWeeks]
        
        print u"\n打印self.recurrWeeks"#FORDEBUG
        print self.recurrWeeks,'\n'#FORDEBUG
    
    def save(self):
        u'''save方法
    生成课程的appointment并储存到Outlook'''
        for i in self.recurrWeeks:
            if i != None:
                self.dataDict['date'] = C.DateDict[int(i + self.rawList[0][0])]
                
                print u"打印['date']"#FORDEBUG
                print self.dataDict['date']#FORDEBUG
                
                #TODO
                #call apptGen(self.dataDict)


def main():
    '''main函数
    主程序'''
    show('welcome')                                                 #显示欢迎信息

    show(0)                                                         #获取
    termFirstDayStr = raw_input()                                   #学期第一周周一日期 --> string
    show(1)                                                         #获取
    termTotalWeeks = int(raw_input())                               #学期总周数 --> integer

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

        
def apptGen(dataDict):
    u'''apptGen()函数
    从传入的dataDict参数中获取具体属性以生成指定的AppointmentItem'''
    import win32com.client                                  #导入必需的模块

    o = win32com.client.Dispatch("Outlook.Application")     #新建一个Outlook.Application实例 o
    a = o.CreateItem(1)                                     #新建一个AppointmentItem实例 a

    a.Start = dataDict['date']+' '+dataDict['startTime']    #设置课程开始时间
    a.Duration = dataDict['duration']                       #设置课程时长
    a.Subject = dataDict['subject']                         #设置课程名称
    a.Location = dataDict['location']                       #设置授课教室

    p = a.GetRecurrencePattern()                            #获取AppointmentItem.GetRecurrencePattern()对象以修改a的重复类型
    p.RecurrenceType = 0                                    #设置重复类型为 0 == olRecursDaily

    p.PatternStartDate = dataDict['date']                   #设置重复开始日期
    p.PatternEndDate = dataDict['date']                     #结束日期为同一天

    a.Save()                                                #储存 a


if __name__ == '__main__':
    C = Const()                                                               #创建Const实例C以提供相应常量
    try:
        main()                                                         #初始化
        show('debug')
        print C.DateDict                                                      #打印weekDayDict
        show('debug')#FORDEBUG

        show(2)
        while True:
            tmp = raw_input()
            if tmp == '':
                break
            else:
                Class(tmp)
    except:
        pass 