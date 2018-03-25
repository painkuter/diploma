import numpy as np

from sklearn.decomposition import PCA

# import sys

# a = np.arange(15).reshape(3, 5)

# =================================================================
N = 6
#z = np.arange(N)  # normal profitability
# R = np.arange(N) # stocks
# R = np.matrix(np.random.randint(1,10,(N, 20)))
R = np.random.randint(1, 10, (N, 20))
print(R)
# sys.stdout.write("\rStocks:\n{}\n".format(R))
# =================================================================

# np.mean() - mathematical expectation







pca = PCA()
pca.fit(R)

print(pca.explained_variance_)

sigma = []
mu = []
z = []
for Ri in R:
    mu_i = np.mean(Ri)
    mu.append(mu_i)

    sigma_i = np.var(Ri)
    sigma.append(sigma_i)
    z.append((Ri-mu_i)/sigma_i)

# while i < N:
#     while j < N:
#         sigma_S += sigma[i]*sigma[j]*np.mean(z[i])
#
#
# print(mu)

# E = np.array(np.mean(R), ndmin=2)
# print(E)


#
# zz = np.multiply(z,z)
#
# print(zz)
#
# sigma_S = sum(z*z)
#
# print(sigma_S)
# print(z[1])
