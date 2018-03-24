import openpyxl as xl
import datetime

path = "data"

from os import listdir
from os.path import isfile, join

files = [f for f in listdir(path)]

# start_date =

# last_date = datetime.datetime(year=2014, month=1, day=1)
last_date = None
last_price = 0

data = dict()  # map

for file in files:

    wb = xl.load_workbook(path + "/" + file)

    # grab the active worksheet
    ws = wb.active
    # create new
    new_ws = wb.create_sheet(title="Draft")

    i = 0
    j = 1
    for row in ws.rows:

        i += 1

        if i > 7:  # ignore header

            if (type(row[1].value) is float) & (type(row[0].value) is datetime.datetime):
                date = row[0].value
                price = str(row[1].value)

                if i == 9:
                    last_date = date
                else:
                    if date != last_date - datetime.timedelta(days=1):

                        delta = (last_date - date).days

                        print("last day: " + str(last_date) + " date: "+ str(date) + "  delta: " + str(delta))
                        for k in range(1, delta):
                            new_date = last_date - datetime.timedelta(days=k)
                            print(new_date)
                            new_ws["A" + str(j)] = new_date.date()
                            new_ws["B" + str(j)] = last_price
                            j += 1

                new_ws["A" + str(j)] = date.date()
                new_ws["B" + str(j)] = price
                j += 1

                last_date = row[0].value
                last_price = row[1].value

    data[file] = i

    wb.save(path + "/" + file)

for key in data.keys():
    print(key + ": " + str(data[key]))
