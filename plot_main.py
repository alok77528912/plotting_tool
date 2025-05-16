import textwrap

import matplotlib.pyplot as plt
import numpy as np
import datetime as dt
import pandas as pd
import os



def myplotext(file_path,sheet_key, section_key,linex, liney,plot_col,x_key,y_key, extra_key, xrange, yrange, limit_label, title, xy:bool=0, x:bool=0, y:bool=0, filelen = 27, file_name='', tx_power:int=0):

    ### Start Fetch files name and case list ###
    include_file_name = file_name.split("&&")
    global df_plot1
    global name
    global plot_rows
    files_list = []
    case_list = []
    case_list_extra_key = []
    for files in os.listdir(file_path):
        if '~$' not in files and files.endswith('.xlsx') or files.endswith('.xls'):
            file_name_filter = 0
            for name_filter in include_file_name:
                if name_filter in files:
                    file_name_filter=file_name_filter+1
                elif '+' in name_filter:
                    or_filter = name_filter.split("+")
                    for orfilter in or_filter:
                        if orfilter in files:
                            file_name_filter = file_name_filter + 1
                            break


            if file_name_filter == len(include_file_name) or len(file_name) == 0:

                file = os.path.join(file_path,files)
                files_list.append(files)
                df = pd.read_excel(file, sheet_name=sheet_key)

                df_unique = df[section_key].drop_duplicates().reset_index()

                for i in range(int(len(df_unique))):
                    uni = []
                    for j in range(len(section_key)):
                        column_name = section_key[j]
                        value = df_unique.loc[i, str(column_name)]
                        uni.append(value)

                    if uni not in case_list:
                        case_list.append(uni)

                if isinstance(extra_key, list):
                    df_unique_extra = df[extra_key].drop_duplicates().reset_index()
                    for i in range(int(len(df_unique_extra))):
                        uni1 = []
                        for j in range(len(extra_key)):
                            column_name = extra_key[j]
                            value = df_unique_extra.loc[i, str(column_name)]
                            uni1.append(value)
                        if uni1 not in case_list_extra_key:
                            case_list_extra_key.append(uni1)

    ### End Fetch files name and case list ###

    #### Start Print files name and case list ###
    if len(files_list) == 0:
        print("There is no file in the path")
        return

    files_list.sort()
    case_list.sort()
    case_list_extra_key.sort()

    print("There are below files in the path:\n")
    for k in files_list:
        print(k)
    print("Case list: ")
    print(case_list)
    #### End Print files name and case list ###

    ### Start calculate plot quantity and columns ###
    plots_qty = len(case_list)
    if plot_col != 0:
        plot_rows = int(plots_qty/plot_col)
        if plots_qty%plot_col != 0:
            plot_rows += 1
    plot_number = 1
    index_n = 1
    ### End calculate plot quantity and columns ###

    ### Start to make pictures ###
    if plot_col != 0:
        plt.figure(figsize=(((15+plot_col) * plot_col), (((8+plot_col) * plot_rows))))

    legend_name = []
    legend_colors = []

    ### Start plots test case one by one ###
    for i in range(0, len(case_list), 1):
        if plot_col == 0:
            plt.figure(figsize=(15.5, 8))
        ### Start print test case ###
        print("Now plot ", end='')
        for_title = []
        for j in range(len(section_key)):
            plot = f"{section_key[j]}={case_list[i][j]}"
            for_title.append(plot)
        print(','.join(for_title))
        ### Ent print test case ###

        legend_names = []
        x_y_values = []

        ### Start to fetch data from files ###
        for files in files_list:

            print("Processing file "+files)
            file = os.path.join(file_path, files)
            df_plot = pd.read_excel(file, sheet_name=sheet_key)
            for j in range(len(section_key)):
                df_plot = df_plot[df_plot[section_key[j]] == case_list[i][j]]
            if extra_key == '':
                if len(list(df_plot[x_key])) != 0:
                    x_values = list(df_plot[x_key])
                    y_values =  list(df_plot[y_key])

                    files = files[:filelen]

                    legend_names.append(f"{files}")
                    x_y_list = [x_values,y_values]
                    x_y_values.append(x_y_list)
                    print(f"[{min(x_values)},{max(x_values)}]", end='\n')
                    print(f"[{min(y_values)},{max(y_values)}]", end='\n')
                else:
                    print("No data found regarding current test case")

        # For extra key

            elif isinstance(extra_key, list):
                for f in range(len(case_list_extra_key)):
                    df_plot1 = df_plot
                    name = ''
                    for fn in range(len(extra_key)):
                        df_plot1 = df_plot1[df_plot1[extra_key[fn]]] == case_list_extra_key[f][fn]
                        name1 = f"{extra_key[fn]} = {case_list_extra_key[f][fn]}, "
                        name = name+name1
                    if len(list(df_plot1[y_key])) != 0:
                        x_values = list(df_plot1[x_key])
                        y_values = list(df_plot1[y_key])
                        files = files[:filelen]
                        if len(files) != 0:
                            legend_name_with_ex_key = f"{name} || {files}"
                        else:
                            legend_name_with_ex_key = f"{name}"
                        legend_names.append(f"{legend_name_with_ex_key}")
                        x_y_list = [x_values, y_values]
                        x_y_values.append(x_y_list)

            else:
                unique_values = df_plot[extra_key].unique()
                for values in unique_values:
                    df_plot_extra_key = df_plot[df_plot[extra_key]==values]
                    x_values = list(df_plot_extra_key[x_key])
                    y_values = list(df_plot_extra_key[y_key])
                    files = files[:filelen]
                    if filelen != 0:
                        legend_update = f"{extra_key}={values}_{files}"
                    else:
                        legend_update = f"{extra_key}={values}"
                    legend_names.append(legend_update)
                    x_y_list = [x_values, y_values]
                    x_y_values.append(x_y_list)
                    print(f"[{min(x_values)},{max(x_values)}]", end='\n')
                    print(f"[{min(y_values)},{max(y_values)}]", end='\n')

        print(f'There are {len(legend_names)} legends')
        ### End to fetch data from files ###

        ### Start plotting Limit ###
        limit_number = len(linex)
        # face_color = (0.90, 0.90, 0.90)
        limit_color = ['red', 'deeppink', 'blue', 'purple', 'brown', 'gray', 'orange']
        if plot_col != 0:
            plt.subplot(plot_rows, int(plot_col), index_n)

        x_axis_max, y_axis_max = -9999999999, -999999999999
        x_axis_min, y_axis_min = 999999999999, 99999999999

        if tx_power == 0:
            for z in range(limit_number):
                if len(limit_label) == 0:
                    limit_l = f'Limit {z}'
                else:
                    limit_l = f"{limit_label[z]}"
                limit_l = '\n'.join(textwrap.wrap(limit_l, width=27))
                plt.plot(linex[z], liney[z], color=limit_color[z], linestyle='dashed', label=limit_l, linewidth=1)

                if y_axis_max < max(liney[z]):
                    y_axis_max = max(liney[z])
                if y_axis_min > min(liney[z]):
                    y_axis_min = min(liney[z])

        ### End plotting Limit ###

        xdiff,ydiff = 0,0
        limit_times = 0

        ### Start Plotting Data ###
        for k in range(len(x_y_values)):
            plot_label_wrap= '\n'.join(textwrap.wrap(legend_names[k],width=31))

            if plot_label_wrap not in legend_name:
                plot_color = (np.random.random(), np.random.random(), np.random.random())
            else:
                plot_color = legend_colors[legend_name.index(plot_label_wrap)]

            if tx_power != 0 and limit_times == 0:
                tx_x_limit = [min(x_y_values[k][0]), max(x_y_values[k][0])]
                tx_y_lower_limit = [min(x_y_values[k][0])+tx_power, max(x_y_values[k][0])+tx_power]
                tx_y_upper_limit = [min(x_y_values[k][0]) - tx_power, max(x_y_values[k][0]) - tx_power]
                plt.plot(tx_x_limit, tx_y_upper_limit, color='red', linestyle='dashed', label="limit", linewidth=1)
                plt.plot(tx_x_limit, tx_y_lower_limit, color='red', linestyle='dashed', linewidth=1)
                limit_times = limit_times+1

                if y_axis_max < tx_y_upper_limit[1]:
                    y_axis_max = tx_y_upper_limit[1]
                if y_axis_min > tx_y_lower_limit[0]:
                    y_axis_min = tx_y_lower_limit[0]
            if '_baseline' in plot_label_wrap:
                plt.plot(x_y_values[k][0], x_y_values[k][1], label = plot_label_wrap, marker = 'o', markersize = 1.5, linestyle = 'dashed', linewidth=0.8)
            else:
                plt.plot(x_y_values[k][0], x_y_values[k][1], label=plot_label_wrap, marker='o', markersize=1.5, linewidth=0.8)
            current_color = plt.gca().lines[-1].get_color()
            if plot_label_wrap not in legend_name:
                legend_name.append(plot_label_wrap)
                legend_colors.append(current_color)

            if x_axis_max < max(x_y_values[k][0]):
                x_axis_max = max(x_y_values[k][0])
            if x_axis_min > min(x_y_values[k][0]):
                x_axis_min = min(x_y_values[k][0])
            if y_axis_max < max(x_y_values[k][1]):
                y_axis_max = max(x_y_values[k][1])
            if y_axis_min > min(x_y_values[k][1]):
                y_axis_min = min(x_y_values[k][1])
            if x_axis_max == x_axis_min:
                xdiff = 1
            elif (x_axis_max-x_axis_min) > 16 and xdiff != 1:
                xdiff = round((int(x_axis_max)-int(x_axis_min))/16)
            elif 16 >= (x_axis_max - x_axis_min) >= 8 and xdiff != 1:
                xdiff = 1
            elif (x_axis_max-x_axis_min) < 8 and xdiff != 1:
                xdiff = round(((x_axis_max - x_axis_min)/10), 1)
            if y_axis_max == y_axis_min:
                ydiff = 1
            elif (y_axis_max - y_axis_min) > 16 and ydiff != 1:
                ydiff = round((int(y_axis_max) - int(y_axis_min)) / 16)
            elif 8 <= (y_axis_max-y_axis_min) <= 16 and ydiff != 1:
                ydiff = 1
            elif (y_axis_max-y_axis_min) < 8 and ydiff != 1:
                ydiff = round(((y_axis_max - y_axis_min) / 10),1)

            ### Start Print coordinates ###
            round_y = []
            if bool(xy):
                for elements in x_y_values[k][1]:
                    round_y.append(round(elements, 2))
                for f in range(len(round_y)):
                    plt.annotate(f"({x_y_values[k][0][f]},{round_y[f]})",(x_y_values[k][0][f], round_y[f]), color=current_color, size=8)
            if bool(x):
                for elements in x_y_values[k][1]:
                    round_y.append(round(elements, 2))
                for f in range(len(round_y)):
                    plt.annotate(f"{x_y_values[k][0][f]}", (x_y_values[k][0][f], x_y_values[k][1][f]), color=current_color, size=8)
            if bool(y):
                for elements in x_y_values[k][1]:
                    round_y.append(round(elements, 2))
                for f in range(len(round_y)):
                    plt.annotate(f"{round_y[f]}", (x_y_values[k][0][f], round_y[f]), color=current_color, size=8)
            ### End Print coordinates ###
        ### End plotting data ###
        if len(legend_names) > 25 and len(plot_label_wrap)<15:
            legend_column = 2
            legend_size = 14
        elif len(legend_names) > 25:
            legend_column = 1
            legend_size = 10
        else:
            legend_column = 1
            legend_size = 14
        if plot_col != 0:
            plt.legend(bbox_to_anchor=(0.994, 1.012), loc='upper left',prop={'size':legend_size}, frameon = True, ncol = legend_column)
        if plot_col == 0:
            plt.legend(bbox_to_anchor=(0.994, 1.012), loc='upper left',prop={'size':legend_size}, frameon = True, ncol = legend_column)

        ### Plot formatting ###

        # plt.gca().spines['bottom'].set_color('white')
        # plt.gca().spines['top'].set_color('white')
        # plt.gca().spines['right'].set_color('white')
        # plt.gca().spines['left'].set_color('white')

        plt.title(','.join(for_title),fontsize=20, y=1.02)
        plt.tick_params(axis='both', labelsize=15)
        plt.xlabel(x_key,fontsize=17)
        plt.ylabel(y_key, fontsize=17)
        plt.grid(True, color='lightgrey', linewidth=0.5)

        # plt.gca().patch.set_facecolor(face_color)

        if len(xrange) != 0:
            plt.xticks(np.arange(xrange[0], xrange[1]+xrange[2], xrange[2]))
            plt.xlim(xrange[0], xrange[1])
        if len(yrange) != 0:
            plt.yticks(np.arange(yrange[0], yrange[1]+yrange[2], yrange[2]))
            plt.ylim(yrange[0], yrange[1])

        if xdiff == 0:
            xdiff = 0.1
        if ydiff == 0:
            ydiff = 0.1

        if len(xrange) == 0 and xdiff >=1:
            plt.xticks(np.arange((round(x_axis_min)-2*xdiff), (round(x_axis_max)+3*xdiff), xdiff))
            plt.xlim((round(x_axis_min)-xdiff), (round(x_axis_max)+xdiff))
        elif len(xrange) == 0:
            plt.xticks(np.arange((round(x_axis_min,1) - 2 * xdiff), (round(x_axis_max,1) + 3 * xdiff), xdiff))
            plt.xlim((round(x_axis_min,1) - xdiff), (round(x_axis_max,1) + xdiff))
        if len(yrange) == 0 and ydiff >=1:
            plt.yticks(np.arange((round(y_axis_min)-2*ydiff), (round(y_axis_max)+3*ydiff), ydiff))
            plt.ylim((round(y_axis_min)-ydiff), (round(y_axis_max)+ydiff))
        elif len(yrange) == 0:
            plt.yticks(np.arange((round(y_axis_min,1) - 2 * ydiff), (round(y_axis_max,1) + 3 * ydiff), ydiff))
            plt.ylim((round(y_axis_min,1) - ydiff), (round(y_axis_max,1) + ydiff))

        #### Start To add date and time in save file ###
        d = dt.datetime.now()
        date = f'{d.strftime("%Y")}{d.strftime("%m")}{d.strftime("%d")}'
        time = f'{d.strftime("%H")}{d.strftime("%M")}{d.strftime("%S")}'
        #### End To add date and time in save file ###

        if plot_col == 0:
            plt.subplots_adjust(right=0.72, left = 0.06, top=0.94, bottom=0.08)
            fig_name = f"{title}_{x_key}_vs_{y_key}_{'_'.join(for_title)}_{date}_{time}.png"
            plt.savefig(os.path.join(file_path, fig_name))
            print("Plot saved : " + fig_name)

        index_n += 1
        plot_number += 1
    ### Configure final plot ###
    if plot_col != 0:
        if len(title) == 0:
            title = legend_names[k]
        plt.suptitle(title, fontsize=20 , fontweight="bold")
        # plt.tight_layout(rect=[0, 0, 1, 0.98])

        if plot_col == 1 and plot_rows == 1:
            plt.subplots_adjust(right=0.68, left=0.06, top=0.88, bottom=0.08)
        elif plot_col == 1 and plot_rows == 2:
            plt.subplots_adjust(right=0.70, left=0.065, top=0.93, bottom=0.039, hspace=0.18, wspace=0.5)
        elif plot_col == 1 and plot_rows == 3:
            plt.subplots_adjust(right=0.70, left=0.061, top=0.93, bottom=0.039, hspace=0.18, wspace=0.5)
        elif plot_col == 1:
            plt.subplots_adjust(right=0.70, left=0.05, top=0.95, bottom=0.033)
        elif plot_rows == 1 and plot_col == 2:
            plt.subplots_adjust(right=0.81, left=0.03, top=0.89, bottom=0.07, hspace=0.2, wspace=0.5)
        elif plot_rows == 2 and plot_col == 2:
            plt.subplots_adjust(right=0.83, left=0.03, top=0.92, bottom=0.05, hspace=0.2, wspace=0.5)
        elif plot_rows == 3 and plot_col == 2:
            plt.subplots_adjust(right=0.83, left=0.03, top=0.95, bottom=0.03, hspace=0.2, wspace=0.5)
        elif plot_rows == 4 and plot_col == 2:
            plt.subplots_adjust(right=0.83, left=0.03, top=0.95, bottom=0.02, hspace=0.2, wspace=0.5)
        elif plot_rows == 5 and plot_col == 2:
            plt.subplots_adjust(right=0.83, left=0.03, top=0.95, bottom=0.02, hspace=0.2, wspace=0.5)
        elif plot_rows == 1 and plot_col == 3:
            plt.subplots_adjust(right=0.87, left=0.02, top=0.89, bottom=0.07, hspace=0.2, wspace=0.5)
        elif plot_rows == 2 and plot_col == 3:
            plt.subplots_adjust(right=0.9, left=0.02, top=0.93, bottom=0.04, hspace=0.2, wspace=0.5)
        else:
            plt.subplots_adjust(hspace=0.2, wspace=0.5)


        tit = f"{title}_{x_key}_vs_{y_key}_{date}_{time}.png"
        plt.savefig(os.path.join(file_path, tit))
        print("Plot saved: " + tit)

