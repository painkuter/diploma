import openpyxl as xl
import datetime

path = "data"
draft = "Draft"
seporator = "====================================================="

from os import listdir
from os.path import isfile, join

files = sorted(listdir(path))

data = dict()  # map

last_date = None
last_price = 0

for file in files:
    print(file)

    wb = xl.load_workbook(path + "/" + file)

    # grab the active worksheet
    ws = wb.active

    # create new or open existed
    try:
        draft_ws = wb[draft]
    except KeyError:
        draft_ws = wb.create_sheet(title=draft)

    i = 0
    j = 1
    for row in ws.rows:


        i += 1

        if i > 8:  # ignore header and first row

            if (type(row[1].value) in [float,  int]) & (type(row[0].value) is datetime.datetime):

                date = row[0].value
                price = row[1].value

                if i == 9:
                    last_date = date

                else:
                    if date != last_date - datetime.timedelta(days=1):

                        delta = (last_date - date).days

                        print("last day: " + str(last_date) + " date: "+ str(date) + "  delta: " + str(delta))
                        for k in range(1, delta):
                            new_date = last_date - datetime.timedelta(days=k)
                            print(str(new_date) + " price: " + str(last_price))
                            draft_ws["A" + str(j)] = new_date.date()
                            draft_ws["B" + str(j)] = last_price
                            j += 1

                draft_ws["A" + str(j)] = date.date()
                draft_ws["B" + str(j)] = price
                j += 1

                last_date = row[0].value
                last_price = row[1].value

    data[file] = j

    wb.save(path + "/" + file)
    print(seporator)

print()

for key in sorted(data.keys()):
    print(key + ": " + str(data[key]))
