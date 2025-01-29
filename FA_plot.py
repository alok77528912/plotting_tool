import numpy as np
import glob
import os
import sys
import pdb

from matplotlib.pyplot import title

cwd = 'D:\Python_Projects\SumPlot'
sys.path.append(cwd)# from MyPlotExt import myplotext
from main import  myplotext
# from main_13_01_2025 import myplotext


# Updated to plot all files inside a folder since 2.0.1
plot_path = 'D:\Python_Projects\SumPlot'

myplotext(file_path=plot_path,sheet_key='BT_RX',section_key=['rate','ant','PktType','DataPattern','PktLength','PktNum','sc_mode','sc_pll_ldo_trim','forceAGC_idx'],\
linex=[[-18,10],[-18,10]], liney=[[0,0],[-1,-1]],plot_col=4,x_key='power',y_key='PER(%)',\
extra_key='',xrange=[-120,3,15],yrange=[],limit_label=['Jaanu','Mood'],title='Alok', y=True)
