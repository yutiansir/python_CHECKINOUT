import pyodbc
import datetime
cnxn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=C:\Users\Killer.B\Documents\VScode\att2000.mdb')
crsr = cnxn.cursor()
print('`````````````` att2000.mdb ``````````````')

def get_range_date():
    start_time = datetime.date(2021,8,10)
    end_time = datetime.date(2028,10,10)
    day_range = list()
    for i in range((end_time - start_time).days + 1):
        day = start_time + datetime.timedelta(days = i)
        day_range.append(str(day))
    return day_range

morning_num = []
afternoon_num = []
for curr_time in get_range_date():
    morning_num.append("INSERT INTO CHECKINOUT values('296', '" + curr_time + " 07:31:18', 'I', '15','1',NULL, '0','AC832105600','1');")
    afternoon_num.append("INSERT INTO CHECKINOUT values('296', '" + curr_time + " 14:26:18', 'I', '15','1',NULL, '0','AC832105600','1');")

print(morning_num)
print(afternoon_num)

for i in morning_num:
    crsr.execute(i)

for k in afternoon_num:
    crsr.execute(k)

crsr.commit()
crsr.close()
cnxn.close()