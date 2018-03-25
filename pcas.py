import openpyxl as xl
import datetime
import numpy as np
from os import listdir
from os.path import isfile, join
from sklearn.decomposition import PCA

start_time = datetime.datetime.now()

max_last_date = datetime.datetime(year=2017, month=12, day=22)
min_first_date = datetime.datetime(year=2017, month=11, day=1)

path = "data"

files = sorted(listdir(path))

R = dict()
i = 0

for file in files:
    print(file)
    wb = xl.load_workbook(path + "/" + file)
    draft_ws = wb["Draft"]

    j = 0
    R[i] = []
    for row in draft_ws.rows:
        #print(row[0].value)
        if min_first_date < row[0].value < max_last_date:
            R[i].append(row[1].value)
            j += 1
    #print(j)
    i += 1

keys = R.keys()

R_ = np.array([R[k] for k in range(0,i)])
#print(R_)

sigma = []
mu = []
z = []
t = 0
for Ri in R_:
    mu_i = np.mean(Ri)
    mu.append(mu_i)

    sigma_i = np.var(Ri)
    sigma.append(sigma_i)

    #d = len(Ri)
    buf = []
    for Rij in Ri:
        buf.append((Rij-mu_i)/sigma_i)
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

for zi in z:
    j = 0
    for zij in zi:
        sigma_S += sigma[j]*sigma[i]*np.mean(np.array(z[i])*np.array(z[j]))
        j += 1
    i += 1

print(sigma)

PCAS_k = 0
k= 1
i =0
for lambda_k in lambda_:
    PCAS_k += sigma[i]/sigma_S*L[k][i]*lambda_[i]

    i += 1
print(PCAS_k)

end_time = datetime.datetime.now()
print()
print(end_time - start_time)
