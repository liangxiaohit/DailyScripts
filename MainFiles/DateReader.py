import datetime
from time import time,localtime,strftime
class Reader:
    will_check_date = False
    # 需要传递到MainFile中
    mon, tue, wed, thu, fri, sat, sun = None, None, None, None, None, None, None
    week_days = [mon, tue, wed, thu, fri, sat, sun]
    def __init__(self):
        print("DateReader is created")

    def creat_date(self) -> str:
        monday = datetime.date.today()
        sunday = datetime.date.today()
        one_day = datetime.timedelta(1)
        while monday.weekday() != 0:
            monday -= one_day
        while sunday.weekday() != 6:
            sunday += one_day

        Reader.week_days[0] = monday
        Reader.week_days[1] = monday + one_day
        Reader.week_days[2] = monday + one_day * 2
        Reader.week_days[3] = monday + one_day * 3
        Reader.week_days[4] = monday + one_day * 4
        Reader.week_days[5] = monday + one_day * 5
        Reader.week_days[6] = sunday

        # 以下是将月份以及年份小于10的在首位补0
        monday_front = str(monday.month)
        monday_end = str(monday.day)
        if monday.month<10:
            monday_front = str(0)+str(monday.month)
        if monday.day<10:
            monday_end = str(0)+str(monday.day)
        sunday_front = str(sunday.month)
        sunday_end = str(sunday.day)
        if sunday.month<10:
            sunday_front = str(0)+str(sunday.month)
        if sunday.day<10:
            sunday_end = str(0)+str(sunday.day)

        path_date = monday_front + monday_end + "-" + sunday_front + sunday_end
        print(path_date)
        return path_date

