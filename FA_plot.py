import numpy as np
import glob
import os
import sys
import pdb

from matplotlib.pyplot import title

cwd = 'D:\Python_Projects\SumPlot'
sys.path.append(cwd)# from MyPlotExt import myplotext
from old.RFPlot_10_04_2025 import  myplotext
# from main_13_01_2025 import myplotext


# Updated to plot all files inside a folder since 2.0.1
plot_path = 'E:\data'

myplotext(file_path=plot_path,sheet_key='11ax_AveragePower',section_key=["protocol","BandWidth"],\
linex=[], liney=[],plot_col=0,x_key='TargetTxPower(dBm)',y_key='AvgPwr(dBm)_stream0',\
extra_key='ChanFreq(MHz)',xrange=[],yrange=[],limit_label=[],title='Alok', y=True)
