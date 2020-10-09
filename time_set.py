import datetime
import time


def time_conv(date_f):
        f_date, t_date = date_f.split('-')
        d, m, y = f_date.split('.')
        d2, m2, y2 = t_date.split('.')
        now = datetime.datetime.now()

        if int(d) > 32 or int(m) > 12 or int(y) < 2019 or int(d) < 0 or int(m) < 0 or int(y) > int(now.year):
            print('Неверная дата')
            return
        if int(d) > int(now.day) and int(m) > int(now.month) and int(y) > int(now.year):
            print('Неверная дата')
            return

        if int(d2) > 32 or int(m2) > 12 or int(y2) < 2019 or int(d2) < 0 or int(m2) < 0 or int(y2) > int(now.year):
            print('Неверная дата')
            return
        if int(d2) > int(now.day) and int(m2) > int(now.month) and int(y2) > int(now.year):
            print('Неверная дата')
            return

        h1 = '00'
        min1 = '00'
        h2 = '23'
        min2 = '59'
        s = '00'

        t1 = datetime.datetime(int(y), int(m), int(d), int(h1), int(min1), int(s))
        t2 = datetime.datetime(int(y2), int(m2), int(d2), int(h2), int(min2), int(s))

        try:

            t1_s1_unix = int(str(time.mktime(t1.timetuple()))[:-2]) + 7200
            t2_s1_unix = int(str(time.mktime(t2.timetuple()))[:-2]) + 7200
            return t1_s1_unix, t2_s1_unix, f_date, t_date

        except Exception as exc:
            print(exc)
            print('Неверный формат даты')


