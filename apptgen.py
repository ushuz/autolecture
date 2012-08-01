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
    print apptGen.__doc__