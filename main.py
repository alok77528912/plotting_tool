import textwrap

import matplotlib.pyplot as plt
import numpy as np
import datetime as dt
import pandas as pd
import os

from matplotlib.pyplot import xticks
from openpyxl.styles.alignment import horizontal_alignments, vertical_aligments

#### Start To add date and time in save file ###

d = dt.datetime.now()
date = f'{d.strftime("%Y")}{d.strftime("%m")}{d.strftime("%d")}'
time = f'{d.strftime("%H")}{d.strftime("%M")}{d.strftime("%S")}'

#### End To add date and time in save file ###

def myplotext(file_path,sheet_key, section_key,linex, liney,plot_col,x_key,y_key, extra_key, xrange, yrange, limit_label, title, xy:bool=0, x:bool=0, y:bool=0):

    ### Start Fetch files name and case list ###

    files_list = []
    case_list = []
    for files in os.listdir(file_path):
        if '~$' not in files and files.endswith('.xlsx') or files.endswith('.xls'):
            file = os.path.join(file_path,files)
            files_list.append(files)
            df = pd.read_excel(file, sheet_name=sheet_key)

            df_unique = df[section_key].drop_duplicates().reset_index()

            if section_key not in case_list:
                case_list.append(section_key)

            for i in range(int(len(df_unique))):
                uni = []
                for j in range(len(section_key)):
                    column_name = section_key[j]
                    value = df_unique.loc[i, str(column_name)]
                    uni.append(value)

                if uni not in case_list:
                    case_list.append(uni)

    ### End Fetch files name and case list ###

    #### Start Print files name and case list ###

    print("There are below files in the path:\n")
    for k in files_list:
        print(k)
    print("Case list: ")
    print(case_list)
    #### End Print files name and case list ###

    ### Start calculate plot quantity and columns ###
    plots_qty = len(case_list)-1
    if plot_col != 0:
        plot_rows = int(plots_qty/plot_col)
        if plots_qty%plot_col != 0:
            plot_rows += 1
    plot_number = 1
    index_n = 1
    ### End calculate plot quantity and columns ###

    ### Start to make pictures ###
    if plot_col != 0:
        plt.figure(figsize=(((18+plot_col) * plot_col), (((8+plot_col) * plot_rows))))


    ### Start plots test case one by one ###
    for i in range(1, len(case_list), 1):
        if plot_col == 0:
            plt.figure(figsize=(18, 8))
        ### Start print test case ###
        print("Now plot ", end='')
        for_title = []
        for j in range(len(section_key)):
            plot = f"{case_list[0][j]}={case_list[i][j]}"
            for_title.append(plot)
        print(','.join(for_title))
        ### Ent print test case ###

        legend_names = []
        x_y_values = []

        ### Start to fetch data from files ###
        for files in os.listdir(file_path):
            if '~$' not in files and files.endswith('.xlsx') or files.endswith('.xls'):
                print("Processing file "+files)
                file = os.path.join(file_path, files)
                df_plot = pd.read_excel(file, sheet_name=sheet_key)
                for j in range(len(section_key)):
                    df_plot = df_plot[df_plot[case_list[0][j]] == case_list[i][j]]
                if extra_key == '':
                    if len(list(df_plot[x_key])) != 0:
                        x_values = list(df_plot[x_key])
                        y_values =  list(df_plot[y_key])
                        legend_names.append(files)
                        x_y_list = [x_values,y_values]
                        x_y_values.append(x_y_list)
                        print(f"[{min(x_values)},{max(x_values)}]", end='\n')
                        print(f"[{min(y_values)},{max(y_values)}]", end='\n')
                    else:
                        print("No data found regarding current test case")

            # For extra key

                else:
                    unique_values = df_plot[extra_key].unique()
                    for values in unique_values:
                        df_plot_extra_key = df_plot[df_plot[extra_key]==values]
                        x_values = list(df_plot_extra_key[x_key])
                        y_values = list(df_plot_extra_key[y_key])
                        legend_update = f"{extra_key}={values}_{files}"
                        legend_names.append(legend_update)
                        x_y_list = [x_values, y_values]
                        x_y_values.append(x_y_list)
                        print(f"[{min(x_values)},{max(x_values)}]", end='\n')
                        print(f"[{min(y_values)},{max(y_values)}]", end='\n')

        print(f'There are {len(legend_names)} legends')
        ### End to fetch data from files ###

        ### Start plotting Limit ###
        limit_number = len(linex)
        face_color = (0.90, 0.90, 0.90)
        limit_color = ['red', 'blue', 'purple', 'brown', 'gray', 'orange']
        if plot_col == 0:
            plt.subplots_adjust(right=0.73, left=0.05, top=0.95, bottom=0.1)
        else:
            plt.subplot(plot_rows, int(plot_col), index_n)
            plt.subplots_adjust(right=0.85, left=0.03, top=0.6, bottom=-0.5, hspace=0.2, wspace=0.1)

        for z in range(limit_number):
            if len(limit_label) == 0:
                limit_l = f'Limit {z}'
                plt.plot(linex[z], liney[z], color=limit_color[z], linestyle='dashed', label=limit_l)
            else:
                limit_l = f"{limit_label[z]}"
                plt.plot(linex[z], liney[z], color=limit_color[z], linestyle='dashed', label=limit_l,linewidth=0.5)
        ### End plotting Limit ###
        x_axis_max,y_axis_max = -1000,-1000
        x_axis_min,y_axis_min = 1000,1000
        xdiff,ydiff = 0,0

        ### Start Plotting Data ###
        for k in range(len(x_y_values)):
            plot_label_wrap= '\n'.join(textwrap.wrap(legend_names[k],width=50))

            plt.plot(x_y_values[k][0], x_y_values[k][1], label=plot_label_wrap,marker='o' , markersize=2.5)
            current_color = plt.gca().lines[-1].get_color()
            if x_axis_max < max(x_y_values[k][0]):
                x_axis_max = max(x_y_values[k][0])
            if x_axis_min > max(x_y_values[k][0]):
                x_axis_min = min(x_y_values[k][0])
            if y_axis_max < max(x_y_values[k][1]):
                y_axis_max = max(x_y_values[k][1])
            if y_axis_min > max(x_y_values[k][1]):
                y_axis_min = min(x_y_values[k][1])
            if x_axis_max == x_axis_min:
                xdiff = 1
            if (x_axis_max-x_axis_min) > 4 and xdiff != 1:
                xdiff = int((int(x_axis_max)-int(x_axis_min))/4)
            if (x_axis_max-x_axis_min) < 4 and xdiff != 1:
                xdiff = (x_axis_max - x_axis_min)/4
            if y_axis_max == y_axis_min:
                ydiff = 1
            if (y_axis_max - y_axis_min) > 4 and ydiff != 1:
                ydiff = int((int(y_axis_max) - int(y_axis_min)) / 4)
            if (y_axis_max-y_axis_min) < 4 and ydiff != 1:
                ydiff = (y_axis_max - y_axis_min) / 4

            ### Start Print coordinates ###
            round_y = []
            for elements in x_y_values[k][1]:
                round_y.append(round(elements,2))
            if bool(xy):
                for f in range(len(round_y)):
                    plt.annotate(f"({x_y_values[k][0][f]},{round_y[f]})",(x_y_values[k][0][f], round_y[f]), color=current_color, size=8)
            if bool(x):
                for f in range(len(round_y)):
                    plt.annotate(f"(x={x_y_values[k][0][f]})", (x_y_values[k][0][f], round_y[f]), color=current_color, size=8)
            if bool(y):
                for f in range(len(round_y)):
                    plt.annotate(f"(y={round_y[f]})", (x_y_values[k][0][f], round_y[f]), color=current_color, size=8)
            ### End Print coordinates ###

        ### End plotting data ###

        if len(xrange) == 0:
            plt.xticks(np.arange((x_axis_min - xdiff), (x_axis_max + xdiff), xdiff))
            plt.xlim((x_axis_min-xdiff), (x_axis_max+xdiff))
        if len(yrange) == 0:
            plt.yticks(np.arange((y_axis_min - ydiff), (y_axis_max + ydiff), ydiff))
            plt.ylim((y_axis_min - ydiff), (y_axis_max + ydiff))

        ### Plot formatting ###
        if plot_col != 0 and plot_number % plot_col == 0:
            plt.legend(bbox_to_anchor=(0.994, 1.012), loc='upper left', facecolor=face_color,prop={'size':20})
        if plot_col == 0:
            plt.legend(bbox_to_anchor=(0.994, 1.012), loc='upper left', facecolor=face_color, prop={'size':20})
        plt.gca().spines['bottom'].set_color('white')
        plt.gca().spines['top'].set_color('white')
        plt.gca().spines['right'].set_color('white')
        plt.gca().spines['left'].set_color('white')
        plt.title(','.join(for_title),fontsize=24)
        plt.xlabel(x_key,fontsize=20)
        plt.ylabel(y_key, fontsize=20)
        plt.grid(True, color='w')
        plt.gca().patch.set_facecolor(face_color)

        if len(xrange) != 0:
            plt.xticks(np.arange(xrange[0], xrange[1], xrange[2]))
            plt.xlim(xrange[0], xrange[1])
        if len(yrange) != 0:
            plt.yticks(np.arange(yrange[0], yrange[1], yrange[2]))
            plt.ylim(yrange[0], yrange[1])

        if plot_col == 0:
            fig_name = f"{x_key}_vs_{y_key}_{'_'.join(for_title)}_{date}_{time}.png"
            plt.savefig(os.path.join(file_path, fig_name))
            print("Plot saved : " + fig_name)

        index_n += 1
        plot_number += 1
    ### Configure final plot ###
    if plot_col != 0:
        if len(title) == 0:
            title = legend_names[k]
        plt.suptitle(title, fontsize=(24 + plot_rows), fontweight="bold")
        plt.tight_layout(rect=[0, 0, 1, 0.98])
        tit = f"{title}_{x_key}_vs_{y_key}_{date}_{time}.png"
        plt.savefig(os.path.join(file_path, tit))
        print("Plot saved: " + tit)