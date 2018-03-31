import openpyxl as xl
import datetime
import numpy as np
import sys
import collections
from os import listdir
from os.path import isfile, join

from scipy.constants import value
from sklearn.decomposition import PCA


def has_duplicates(list_of_values):
    key_map = collections.defaultdict(int)
    for item in list_of_values:
        key_map[item] += 1

    for key in key_map.keys():
        if key_map[key] > 1:
            print("Key: " + str(key) + " Value: " + str(key_map[key]))
            return key
    return -1


start_time = datetime.datetime.now()

max_last_date = datetime.datetime(year=2017, month=12, day=23)
min_first_date = datetime.datetime(year=2017, month=9, day=8)

path = "data"

files = sorted(listdir(path))
N = len(files)
R = dict()
i = 0

for file in files:
    wb = xl.load_workbook(path + "/" + file)
    draft_ws = wb["Draft"]

    j = 0
    R[i] = []
    for row in draft_ws.rows:
        # print(row[0].value)
        if min_first_date < row[0].value < max_last_date:
            R[i].append(float(row[1].value))
            j += 1

    print(str(i) + ": len = " + str(len(R[i])) + " : " + file)
    i += 1


Len = len(R[0])

working_days = []
# checking for holiday
for i in range(0,Len):
    is_holiday = True
    for j in range(0,N):
        if R[j][i] != 0:
            is_holiday = False
    if not is_holiday:
        working_days.append(i)


arr = []
R_buf = []
for i in range(0,Len):
    for j in range(0,N):
        R_buf.append([])
        for day in working_days:
            if i == day:
                R_buf[j].append(R[j][day])
print("Working days:", len(R_buf[0]))


# Checking array's lengths
lenR = len(R[0])
for key in R.keys():
    if len(R[key]) != lenR:
        print("Error in array #" + str(key))
        print(R[key])
        print(len(R[key]))
        setRi = set(R[key])


R_ = np.array([R_buf[k] for k in range(0, N)])
# print(R_)

sigma = []
mu = []
z = []
t = 0
for Ri in R_:
    mu_i = np.mean(Ri)
    mu.append(mu_i)

    sigma_i = np.var(Ri)
    sigma.append(sigma_i)

    # d = len(Ri)
    buf = []
    for Rij in Ri:
        buf.append((Rij - mu_i) / sigma_i)
    z.append(buf)
    t += 1

pca = PCA()
pca.fit(z)

ksi = pca.explained_variance_
L = pca.components_

lambda_ = ksi * ksi
N = len(ksi)

sigma_S = 0
i = 0

print("z: " + str(len(z)))
print("sigma: " + str(len(sigma)))

for zi in z:
    j = 0

    for zj in z:

        E = np.array(zi) * np.array(zj)
        mean = np.mean(E)
        sigma_S += np.sqrt(sigma[j] * sigma[i]) * mean
        j += 1
    i += 1

print(sigma)

PCAS = []
k = 1
i = 0

# print(len(lambda_))

for i in range(0, N):
    PCAS.append(0)
    for k in range(0, N):
        PCAS[i] += sigma[i] / sigma_S * L[k][i] * L[k][i] * lambda_[k]

print(PCAS)

end_time = datetime.datetime.now()
print()
print(end_time - start_time)
