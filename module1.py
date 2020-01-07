import datetime

class Date_collection:
    __days = [0, 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    try:
        years, start_month, end_month = map(int, input('년도, 시작 월, 종료 월을 입력해주세요(ex: 2019 1 3) > ').split())
    except Exception as e:
        print("잘못 입력 하셨습니다.", e)

    if years % 4 == 0:
        __days[2] = 29
    else:
        __days[2] = 28

    __day = __days[start_month]

    def box_length(self):
        sum = 0
        total = 0
        for i in range(1, (Date_collection.end_month - Date_collection.start_month) + 1):
            day2 = Date_collection.__days[Date_collection.start_month + i]
            sum += day2
        total = Date_collection.__day + sum
        return total

    def date_insert(self):
        sum = 0
        for i in range(1, (Date_collection.end_month - Date_collection.start_month)+1):
            day2 = Date_collection.__days[Date_collection.start_month + i]
            sum += day2
        total = Date_collection.__day + sum
        day_count = 1
        a = []
        for j in range(0, total+(Date_collection.end_month - Date_collection.start_month)):
            if day_count <= Date_collection.__days[Date_collection.start_month]:
                a.append(datetime.datetime(Date_collection.years, Date_collection.start_month, day_count).strftime('%Y-%m-%d'))
                day_count = day_count + 1
            elif day_count > Date_collection.__days[Date_collection.start_month]:
                if Date_collection.__days[Date_collection.start_month] != 13:
                    Date_collection.start_month = Date_collection.start_month + 1
                    day_count = 1
                else:
                    Date_collection.years = Date_collection.years + 1
                    Date_collection.start_month = 1
                    day_count = 1
        return a
