import datetime
class Const:
    u'''Const类 代表公用常量
    为程序提供所需常量'''
    ClassRoomBuilding = (u'',u'一教',u'二教',u'三教',u'四教',u'体育场')                                        #授课教学楼
    ClassBeginTime = {'12':'8:00','34':'10:10','56':'13:30','78':'15:40','910':'19:00'}                        #课程开始时间

    def genDateDict(self,termFirstDayStr,termTotalWeeks):
        u'''Const.genDateDict方法
    由传入的termFirstDayStr和termTotalWeeksInt参数\n生成以(AB(A:digit,周目;B:digit,日目))为键的工作日词典dateDict'''''
        termFirstDay = datetime.date(int(termFirstDayStr[:4]),int(termFirstDayStr[4:6]),int(termFirstDayStr[6:]))
    
        weekDayList = []
        weekNo = 11
        weekDayDict = {}
    
        #生成工作日(weekday)日期列表(list)
        for i in range(0,termTotalWeeks*7):
            if not((i-5)%7 == 0 or (i-6)%7 == 0):
                weekDayList += [str(termFirstDay + datetime.timedelta(i))]

        #生成工作日(weekday)日期词典(dict)，以(AB(A:周目;B:日目))为键(key)
        for i in range(1,termTotalWeeks*5+1):
            weekDayDict[weekNo] = weekDayList[i-1]
            if i % 5 == 0:
                weekNo += 6
            else:
                weekNo += 1
        self.DateDict = weekDayDict

if __name__ == '__main__':
    print Const.__doc__