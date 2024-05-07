import numpy as np
from scipy.signal import welch
import matplotlib.pyplot as plt

import matplotlib.cm as cm
import matplotlib.tri as mtri
from matplotlib.axis import rcParams
from mpl_toolkits.mplot3d import Axes3D
from matplotlib.colors import Normalize
from matplotlib.ticker import MultipleLocator, ScalarFormatter

def interpolator(coords, val):
    """Функция интерполяции"""
    return scipy.interpolate.RBFInterpolator(coords, val, kernel='cubic')

hl = 1 + 2
dl = 1 + 8
bl = 1 + 8

shirina = 2 * hl + dl
visota = 2 * hl + bl


fig = plt.figure(figsize=(12, 12), dpi= 80)
grid = plt.GridSpec(3, 3, hspace=0.5, wspace=0.5, width_ratios=[hl/shirina, dl/shirina, hl/shirina], height_ratios=[hl/shirina, bl/shirina, hl/shirina])



x_main = np.arange(1, dl, 1)
y_main = np.arange(1, bl, 1)

x_lr = np.arange(1, hl, 1)
y_lr = np.arange(1, bl, 1)

x_tb = np.arange(1, hl, 1)
y_tb = np.arange(1, dl, 1)

xm, ym = np.meshgrid(x_main, y_main)
xlr, ylr = np.meshgrid(x_lr, y_lr)
xtb, ytb = np.meshgrid(y_tb, x_tb)


a1 = fig.add_subplot(grid[0,1], xticklabels=[], yticklabels=[], xticks=[], yticks=[])  # ВВерх
a1.scatter(xtb, ytb)

a2 = fig.add_subplot(grid[1,0], xticklabels=[], yticklabels=[], xticks=[], yticks=[])  # Лево
a2.scatter(xlr, ylr)

a3 = fig.add_subplot(grid[1,1], xticklabels=[], yticklabels=[], xticks=[], yticks=[])  # Центр
a3.scatter(xm, ym)

a4 = fig.add_subplot(grid[1,2], xticklabels=[], yticklabels=[], xticks=[], yticks=[])  # Право
a4.scatter(xlr, ylr)

a5 = fig.add_subplot(grid[2,1], xticklabels=[], yticklabels=[], xticks=[], yticks=[])  # Низ
a5.scatter(xtb, ytb)


plt.show()
plt.close()