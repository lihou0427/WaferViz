# WaferMapApp b5.1 # 10/23/2021 #
# WaferMapApp b5.4 # 12/05/2021 #

#a=0.1, b=0.9

import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.filedialog import askopenfilename
from PIL import ImageTk, Image

import matplotlib
from matplotlib import gridspec
matplotlib.use('TkAgg')
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import axes3d
from matplotlib import cm
from scipy import stats

import math
import xlsxwriter
import os

from scipy.interpolate import griddata
from matplotlib.patches import Circle
from matplotlib.figure import Figure
from datetime import datetime
import time

plt.rcParams['font.family'] ='Arial'
plt.rcParams['font.size'] =5
plt.rcParams['axes.linewidth'] =0.2
plt.rcParams['figure.dpi'] =250

class app_gui:
    def __init__(self):
        self.window_main = tk.Tk()
        self.window_main.title("WaferViz 5.5")
        self.window_main.geometry("937x530+100+20")
        self.window_main.resizable(1, 1)
        self.window_main.minsize(937, 530)
        self.window_main.columnconfigure(0, weight=1)
        self.window_main.rowconfigure(1, weight=1)

        self.frame_control = tk.Frame(self.window_main, bd=0)
        self.frame_control.grid(row=0, column=0, sticky="nesw", padx=1, pady=0)
        self.frame_control.columnconfigure(0, weight=0)
        self.frame_control.rowconfigure(0, weight=0)

        self.frame_display = tk.Frame(self.window_main, bd=0)
        self.frame_display.grid(row=1, column=0, sticky="nesw", padx=1, pady=0)
        self.frame_display.columnconfigure(0, weight=1)
        self.frame_display.rowconfigure(0, weight=1)
        
        self.button_open = tk.Button(self.frame_control, text="Load", 
                            command=self.open_file, width=10)     
        self.button_open.grid(row=0, column=0, sticky="w", columnspan=2, padx=4,
                            pady=0)
        self.file_path = tk.Label(self.frame_control, text="", justify='left', width=15,
                            anchor="w")
        self.file_path.grid(row=1, column=0, columnspan=4, sticky="w", padx=4,
                            pady=0)    

        self.label_id = tk.Label(self.frame_control, text="Enter ID")
        self.label_id.grid(row=2, column=0, sticky='e', padx=4, pady=0)
        self.entry_id = tk.Entry(self.frame_control, width=10)
        self.entry_id.insert(0, ' Run_101')
        self.entry_id.grid(row=2, column=1, sticky='w', padx=0, pady=0)

        self.radio_sigma = tk.IntVar()
        self.radio_sigma.set(1)
        self.button_sigma = tk.Radiobutton(self.frame_control, text='Data Filter \u03C3',
                            variable=self.radio_sigma, value=1, indicator=0, width=12)
        self.button_sigma.grid(column=8, row=0, columnspan=1, sticky='e', padx=4,
                            pady=0)     
        self.entry_sigma = tk.Entry(self.frame_control, width=6, justify='center')
        self.entry_sigma.insert(0, 6)
        self.entry_sigma.grid(row=0, column=9, sticky='w', padx=4, pady=0)
        
        self.button_outlier = tk.Radiobutton(self.frame_control, text='Mask Outliers', 
                            variable=self.radio_sigma, value=2, indicator=0, width=12)  
        self.button_outlier.grid(column=8, row=1, columnspan=1, sticky='e', padx=4,
                            pady=0)
        self.entry_outlier = tk.Entry(self.frame_control, width=6, justify='center')
        self.entry_outlier.insert(0, 0)
        self.entry_outlier.grid(row=1, column=9, sticky='w', padx=4, pady=0)
     
        self.spacer1 = tk.Label(self.frame_control, width=4, text="")        
        self.spacer1.grid(row=0, column=2, sticky='ew')    
        
        option_list1 = [" T1=C", " T1=D", " T1=E", " T1=F", " T1=G", " T1=H", " T1=I",
                            " T1=J", " T1=K", " T1=L", " T1=M", " T1=N", " T1=O", " T1=P",
                            " T1=Q", " T1=R", " T1=S", " T1=T", " T1=U", " T1=V"]
        self.variable1 = tk.StringVar()
        self.variable1.set(" T1=C")
        self.option_button1 = tk.OptionMenu(self.frame_control, self.variable1,
                            *option_list1)
        self.option_button1.config(width=4)
        self.option_button1.grid(column=4, row=0,  sticky='w', padx=1, pady=0)

        option_list2 = [" T2=0", " T2=C", " T2=D", " T2=E", " T2=F", " T2=G", " T2=H",
                            " T2=I", " T2=J", " T2=K", " T2=L", " T2=M", " T2=N", " T2=O",
                            " T2=P", " T2=Q", " T2=R", " T2=S", " T2=T", " T2=U", " T2=V"]
        self.variable2 = tk.StringVar()
        self.variable2.set(" T2=D")
        self.option_button2 = tk.OptionMenu(self.frame_control, self.variable2,
                            *option_list2)
        self.option_button2.config(width=4)
        self.option_button2.grid(column=5, row=0,  sticky='w', padx=1, pady=0)

        self.radio_value_1 = tk.IntVar()
        self.radio_value_1.set(1)
        self.radio_button_1 = tk.Radiobutton(self.frame_control, text='(T1-T2) / t',
                            variable=self.radio_value_1, value=1, indicator=0, width=10)  
        self.radio_button_1.grid(column=4, row=2, columnspan=2, sticky='w', padx=4,
                            pady=0)
        self.radio_button_2 = tk.Radiobutton(self.frame_control, text='T1-T2', 
                            variable=self.radio_value_1, value=2, indicator=0, width=10)  
        self.radio_button_2.grid(column=4, row=1, columnspan=2, sticky='w', padx=4,
                            pady=0)

        self.entry_run_time = tk.Entry(self.frame_control, width=6, justify='center')
        self.entry_run_time.insert(0, 60)
        self.entry_run_time.grid(row=2, column=6, sticky='w', padx=0, pady=0)
        self.label_run_time = tk.Label(self.frame_control, text="Time")
        self.label_run_time.grid(row=2, column=5, sticky='e', padx=4, pady=0)

        self.spacer2 = tk.Label(self.frame_control, width=3, text="")        
        self.spacer2.grid(row=0, column=7, sticky='ew')

        self.entry_contour = tk.Entry(self.frame_control, width=6, justify='center')
        self.entry_contour.insert(0, 10)
        self.entry_contour.grid(row=0, column=12, sticky='w', padx=2, pady=0)
        self.label_contour = tk.Label(self.frame_control, width=8, text="Contours",
                            anchor="e")
        self.label_contour.grid(row=0, column=11, sticky='e', padx=2, pady=0)

        option_list3 = [" Sign", " Dot", " Value", " None"]
        self.variable3 = tk.StringVar()
        self.variable3.set(" Sign")
        self.option_button3 = tk.OptionMenu(self.frame_control, self.variable3,
                            *option_list3)
        self.option_button3.config(width=4)
        self.option_button3.grid(column=12, row=1,  sticky='w', padx=1, pady=0)
        self.marker = tk.Label(self.frame_control, width=8, text="Markers",
                            anchor="e")        
        self.marker.grid(row=1, column=11, sticky='e', padx=2, pady=0)

        self.spacer3 = tk.Label(self.frame_control, width=3, text="")        
        self.spacer3.grid(row=1, column=10, sticky='w')
        self.spacer4 = tk.Label(self.frame_control, width=1, text="")        
        self.spacer4.grid(row=0, column=13, sticky='ew')
        self.spacer5 = tk.Label(self.frame_control, width=3, text="")        
        self.spacer5.grid(row=0, column=16, sticky='ew')

        self.var_limits = tk.IntVar()
        self.check_widget = tk.Checkbutton(self.frame_control, text='Set Range',
                            variable=self.var_limits)
        self.check_widget.grid(row=2, column=11, sticky='w', padx=0, pady=0,
                            columnspan=2)        
        self.limits_entry = tk.Entry(self.frame_control, width=6, justify='center')
        self.limits_entry.insert(0, '1000')
        self.limits_entry.grid(row=2, column=12, sticky='e', padx=4, pady=0)
        
        self.rotation_entry = tk.Entry(self.frame_control, width=6, justify='center')
        self.rotation_entry.insert(0, 0)
        self.rotation_entry.grid(row=0, column=15, sticky='w', padx=4, pady=0)
        self.label_rotation = tk.Label(self.frame_control, width=10, text="Set Rotation",
                            anchor="e")
        self.label_rotation.grid(row=0, column=14, sticky='e', padx=2, pady=0)
    
        self.button_run = tk.Button(self.frame_control, text="Run", 
                            command=self.plot_graph, width=10)
        self.button_run.grid(row=0, column=17, sticky="e", padx=4, pady=0)

        self.button_exit = tk.Button(self.frame_control, text="Save",
                            command=self.save_file, width=10)
        self.button_exit.grid(row=1, column=17, sticky="e", padx=4, pady=0)  
        
        self.button_exit = tk.Button(self.frame_control, text="Exit",
                            command=self.window_main.destroy, width=10)
        self.button_exit.grid(row=2, column=17, sticky="e", padx=4, pady=0)
  
        self.canvas_graph = tk.Canvas(self.frame_display, bg='white', bd=0,
                            scrollregion=[0,0,2000,2000])
        self.canvas_graph.grid(row=0, column=0, sticky="nesw")
        self.canvas_graph.columnconfigure(0, weight=1)
        self.canvas_graph.rowconfigure(0, weight=1)

        self.scrollbar_vertical = tk.Scrollbar(self.frame_display, orient="vertical",
                            relief=tk.SUNKEN, bd=0, width=16)
        self.scrollbar_vertical.grid(row=0, column=1, sticky="ns")
        self.canvas_graph.configure(yscrollcommand=self.scrollbar_vertical.set)
        self.scrollbar_vertical.config(command=self.canvas_graph.yview)

        self.scrollbar_horizontal = tk.Scrollbar(self.frame_display, orient="horizontal",
                            relief=tk.SUNKEN, bd=0, width=16)
        self.scrollbar_horizontal.grid(row=1, column=0, sticky="ew")
        self.canvas_graph.configure(xscrollcommand=self.scrollbar_horizontal.set)
        self.scrollbar_horizontal.config(command=self.canvas_graph.xview)
        
        tk.mainloop()

    def open_file(self):
        global filepath
        global df_data
        filepath = askopenfilename(filetypes=[("Excel", "*.xlsx"), ("Excel", "*.xls"),
                            ("csv", "*.csv"), ("All", "*.*")])
        path = filepath
        self.file_path.configure(text=path)
        
        if filepath.endswith == ".csv":
            df_data = pd.read_csv(filepath)
        else:
            df_data = pd.read_excel(filepath)
        pass

        tk.mainloop()

    def plot_graph(self):
        self.canvas_graph.delete('all')       
        id = self.entry_id.get()
              
        col_size = len(df_data.columns)
        if (col_size == 3):
            df_data.columns = ['A', 'B', 'C']
        elif (col_size == 4):
            df_data.columns = ['A', 'B', 'C', 'D']
        elif (col_size == 5):
            df_data.columns = ['A', 'B', 'C', 'D', 'E']
        elif (col_size == 6):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F']
        elif (col_size == 7):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
        elif (col_size == 8):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        elif (col_size == 9):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        elif (col_size == 10):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        elif (col_size == 11):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
        elif (col_size == 12):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
        elif (col_size == 13):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        elif (col_size == 14):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
        elif (col_size == 15):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
        elif (col_size == 16):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P']
        elif (col_size == 17):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q']
        elif (col_size == 18):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R']
        elif (col_size == 19):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S']      
        elif (col_size == 20):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T']
        elif (col_size == 21):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T', 'U']
        elif (col_size == 22):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T', 'U', 'V']
        else:
            pass

        var1 = self.variable1.get()
        if (var1 == " T1=C"):
            z1 = df_data["C"]
        elif (var1 == " T1=D"):
            z1 = df_data["D"]
        elif (var1 == " T1=E"):
            z1 = df_data["E"]
        elif (var1 == " T1=F"):
            z1 = df_data["F"]
        elif (var1 == " T1=G"):
            z1 = df_data["G"]
        elif (var1 == " T1=H"):
            z1 = df_data["H"]
        elif (var1 == " T1=I"):
            z1 = df_data["I"]
        elif (var1 == " T1=J"):
            z1 = df_data["J"]
        elif (var1 == " T1=K"):
            z1 = df_data["K"]
        elif (var1 == " T1=L"):
            z1 = df_data["L"]
        elif (var1 == " T1=M"):
            z1 = df_data["M"]
        elif (var1 == " T1=N"):
            z1 = df_data["N"]
        elif (var1 == " T1=O"):
            z1 = df_data["O"]
        elif (var1 == " T1=P"):
            z1 = df_data["P"]
        elif (var1 == " T1=Q"):
            z1 = df_data["Q"]
        elif (var1 == " T1=R"):
            z1 = df_data["R"]
        elif (var1 == " T1=S"):
            z1 = df_data["S"]
        elif (var1 == " T1=T"):
            z1 = df_data["T"]
        elif (var1 == " T1=U"):
            z1 = df_data["U"]
        elif (var1 == " T1=V"):
            z1 = df_data["V"]    
        else:
            pass

        var2 = self.variable2.get()
        if (var2 == " T2=0"):
            z2 = 0
        elif (var2 == " T2=C"):
            z2 = df_data["C"]
        elif (var2 == " T2=D"):
            z2 = df_data["D"]
        elif (var2 == " T2=E"):
            z2 = df_data["E"]
        elif (var2 == " T2=F"):
            z2 = df_data["F"]
        elif (var2 == " T2=G"):
            z2 = df_data["G"]
        elif (var2 == " T2=H"):
            z2 = df_data["H"]
        elif (var2 == " T2=I"):
            z2 = df_data["I"]
        elif (var2 == " T2=J"):
            z2 = df_data["J"]
        elif (var2 == " T2=K"):
            z2 = df_data["K"]
        elif (var2 == " T2=L"):
            z2 = df_data["L"]
        elif (var2 == " T2=M"):
            z2 = df_data["E"]
        elif (var2 == " T2=N"):
            z2 = df_data["F"]
        elif (var2 == " T2=O"):
            z2 = df_data["O"]
        elif (var2 == " T2=P"):
            z2 = df_data["P"]
        elif (var2 == " T2=Q"):
            z2 = df_data["Q"]
        elif (var2 == " T2=R"):
            z2 = df_data["R"]
        elif (var2 == " T2=S"):
            z2 = df_data["S"]
        elif (var2 == " T2=T"):
            z2 = df_data["T"]
        elif (var2 == " T2=U"):
            z2 = df_data["U"]
        elif (var2 == " T2=V"):
            z2 = df_data["V"]
        else:
            pass

# data filter
        z = z1 - z2
        data = [df_data["A"], df_data["B"], z]
        headers = ["A", "B", "Z"]
        orignal_df = pd.concat(data, axis=1, keys=headers)
                                    
        value = self.radio_sigma.get()
        if (value == 1):
            sigma_var = float(self.entry_sigma.get())
            z_scores = np.abs(stats.zscore(orignal_df))
            df = orignal_df[(z_scores < sigma_var).all(axis=1)]
            N1 = str(orignal_df.shape[0] - df.shape[0])
        else:
            N2 = int(self.entry_outlier.get())
            orignal_df["dist"] = np.abs(orignal_df["Z"]  - np.mean(orignal_df["Z"]))
            sort_df = orignal_df.sort_values(by="dist", ascending=False)
            df = sort_df.iloc[N2:]
            Z = df["Z"]
            stdev = np.std(Z)
            top = max(df["dist"])
            sigma2 = round(top/stdev, 1)

# contour
        value_rotation = float(self.rotation_entry.get())            
        if (value_rotation == 0):

            x = df["A"]
            y = df["B"]
            z = df["Z"] 

            fig1 = plt.figure(figsize=(2.5, 2.5))
            spec1 = gridspec.GridSpec(ncols=1, nrows=1)
            ax1 = fig1.add_subplot(spec1[0, 0])
            n = 700

            x_grid = np.linspace(np.min(x), np.max(x), n)
            y_grid = np.linspace(np.min(y), np.max(y), n)
            X, Y = np.meshgrid(x_grid, y_grid)
            Z = griddata((x, y), z, (X, Y), method="cubic")
            Ave = round(np.mean(z), 1)
        
            value1 = self.radio_value_1.get()
            if (value1 == 2):
                pass
            else:
                time_pre = self.entry_run_time.get()
                if (time_pre):
                    time = float(self.entry_run_time.get())
                    Rate_pre = abs(Ave)*(60/time)
                    Rate = round(Rate_pre)
                else:
                    pass
            Std = np.std(z)
            Std_percent = round(100*Std/abs(Ave), 2)
            Max = round(max(z))
            Min = round(min(z))
            Nonu = round(0.5*100*(Max-Min)/abs(Ave), 2)
            Range = round(max(z) - min(z))

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_bt_pre = self.limits_entry.get()
                limits_bt = abs(float(limits_bt_pre))
                limits_b = Ave - limits_bt/2
                limits_t = Ave + limits_bt/2
                num = int(self.entry_contour.get())
                if (num >100):
                    num = 100
                    levels = np.linspace(limits_b, limits_t, num)
                else:
                    levels = np.linspace(limits_b, limits_t, num)
                cp = plt.contourf(X, Y, Z, levels = levels, cmap=plt.cm.turbo, alpha=0.95)
                fig1.colorbar(cp, ax=ax1, shrink=0.5, orientation="vertical")

            else:
                contour = int(self.entry_contour.get()) 
                if (contour >100):
                    contour = 100
                else:
                    pass
                cp = plt.contourf(X, Y, Z, contour, cmap=plt.cm.turbo, alpha=0.95)
                cbar1 = fig1.colorbar(cp, ax=ax1, shrink=0.5, orientation="vertical")
                cbar1.ax.locator_params(nbins=6)

            if (value1 == 2):
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'NonU %)')
                plt.text(-75,-270, 'StdD %')
                plt.text(-75,-285, 'Max')
                plt.text(-75,-300, 'Min')
                plt.text(-75,-315, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Nonu)                                      
                plt.text(45,-270, Std_percent)
                plt.text(45,-285, Max)
                plt.text(45,-300, Min)
                plt.text(45,-315, Range)
            else:  
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'Rate (Å/min)')
                plt.text(-75,-270, 'NonU %')
                plt.text(-75,-285, 'StdD %')
                plt.text(-75,-300, 'Max')
                plt.text(-75,-315, 'Min')
                plt.text(-75,-330, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Rate)                                      
                plt.text(45,-270, Nonu)
                plt.text(45,-285, Std_percent)
                plt.text(45,-300, Max)
                plt.text(45,-315, Min)
                plt.text(45,-330, Range)
            
            value = self.radio_sigma.get()
            if (value == 1):
                plt.text(-75,-345, 'Points removed')
                plt.text(45,-345, N1)
            else:
                plt.text(-75,-345, 'Sigma:')
                plt.text(45,-345, sigma2)
        
            circ = Circle((0, 0), 150, facecolor='None', edgecolor='black', lw=0.2,
                              alpha=0.001)
            ax1.add_patch(circ)
            ax1.set_aspect('equal', adjustable='box')

            var3 = self.variable3.get()
            if (var3 == " Value"):
                for x, y ,z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    plt.annotate(label, (x, y), textcoords="offset points", xytext=(0, -3),
                             ha='center', fontsize=3.8, alpha=0.75)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            elif (var3 == " Dot"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.3)
            elif (var3 == " None"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            else:
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    if (z > Ave):
                        ax1.scatter(x, y, marker='+', s=13, linewidths=0.3, color='k', alpha=0.7)
                    else:
                        ax1.scatter(x, y, marker='_', s=13, linewidths=0.3, color='k', alpha=0.7)

            ax1.set_xlabel('X (mm)')
            ax1.set_ylabel('Y (mm)')    
            ax1.set_title(id)
    
            fig1.tight_layout(pad=0.1)
            plot_id1 = "contour"
            plt.savefig(plot_id1, bbox_inches='tight')        
            img1 = Image.open("contour.png")
            plot_id1_resized = img1.resize((340, 370))
            plot_id1_resized.save("contour_resized.png")     
            img11 = ImageTk.PhotoImage(Image.open("contour_resized.png"))
            self.canvas_graph.create_image(0, 35, anchor="nw", image=img11)
            os.remove("contour_resized.png")
            plt.clf()
            plt.close(fig1)

        else:            
            fig4 = plt.figure(figsize=(2.5, 2.5))
            spec4 = gridspec.GridSpec(ncols=1, nrows=1)
            ax5 = fig4.add_subplot(spec4[0, 0])
            radian_rotation = math.radians(value_rotation)

            x_orig = df["A"]
            y_orig = df["B"]        
            x = math.cos(radian_rotation)*x_orig + math.sin(radian_rotation)*y_orig
            y = - math.sin(radian_rotation)*x_orig + math.cos(radian_rotation)*y_orig
            z = df["Z"]
        
            n = 700 
            x_grid = np.linspace(np.min(x), np.max(x), n)
            y_grid = np.linspace(np.min(y), np.max(y), n)
            X, Y = np.meshgrid(x_grid, y_grid)
            Z = griddata((x, y), z, (X, Y), method="cubic")
            Ave = round(np.mean(z), 1)
                    
            value1 = self.radio_value_1.get()
            if (value1 == 2):
                pass
            else:
                time_pre = self.entry_run_time.get()
                if (time_pre):
                    time = float(self.entry_run_time.get())
                    Rate = round(abs(Ave)*(60/time))
                else:
                    pass
            
            Std = np.std(z)
            Std_percent = round(100*Std/abs(Ave), 2)
            Max = round(max(z))
            Min = round(min(z))
            Nonu = round(0.5*100*(Max-Min)/abs(Ave), 2)
            Range = round(max(z) - min(z))

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_bt_pre = self.limits_entry.get()
                limits_bt = abs(float(limits_bt_pre))
                limits_b = Ave - limits_bt/2
                limits_t = Ave + limits_bt/2
                num = int(self.entry_contour.get())
                levels = np.linspace(limits_b, limits_t, num)
                cp = plt.contourf(X, Y, Z, levels = levels, cmap=plt.cm.turbo, alpha=0.95)
                fig4.colorbar(cp, ax=ax5, shrink=0.5, orientation="vertical")

            else:
                contour = int(self.entry_contour.get())
                cp = plt.contourf(X, Y, Z, contour, cmap=plt.cm.turbo, alpha=0.95)
                cbar2 = fig4.colorbar(cp, ax=ax5, shrink=0.5, orientation="vertical")
                cbar2.ax.locator_params(nbins=6)
                
            ax5.set_xlabel('X (mm)')
            ax5.set_ylabel('Y (mm)')    

            if (value1 == 2):
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'NonU %)')
                plt.text(-75,-270, 'StdD %')
                plt.text(-75,-285, 'Max')
                plt.text(-75,-300, 'Min')
                plt.text(-75,-315, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Nonu)                                      
                plt.text(45,-270, Std_percent)
                plt.text(45,-285, Max)
                plt.text(45,-300, Min)
                plt.text(45,-315, Range)
            else:  
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'Rate (Å/min)')
                plt.text(-75,-270, 'NonU %')
                plt.text(-75,-285, 'StdD %')
                plt.text(-75,-300, 'Max')
                plt.text(-75,-315, 'Min')
                plt.text(-75,-330, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Rate)                                      
                plt.text(45,-270, Nonu)
                plt.text(45,-285, Std_percent)
                plt.text(45,-300, Max)
                plt.text(45,-315, Min)
                plt.text(45,-330, Range)

            value = self.radio_sigma.get()
            if (value == 1):
                plt.text(-75,-345, 'Points removed')
                plt.text(45,-345, N1)
            else:
                plt.text(-75,-345, 'Sigma:')
                plt.text(45,-345, sigma2)
                        
            circ = Circle((0, 0), 150, facecolor='None', edgecolor='black', lw=0.2,
                          alpha=0.0001)
            ax5.add_patch(circ)
            ax5.set_aspect('equal', adjustable='box')

            var3 = self.variable3.get()
            if (var3 == " Value"):
                for x, y ,z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    plt.annotate(label, (x, y), textcoords="offset points", xytext=(0, -3),
                             ha='center', fontsize=3.8, alpha=0.75)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            elif (var3 == " Dot"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.3)
            elif (var3 == " None"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            else:
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    if (z > Ave):
                        ax5.scatter(x, y, marker='+', s=13, linewidths=0.3, color='k', alpha=0.7)
                    else:
                        ax5.scatter(x, y, marker='_', s=13, linewidths=0.3, color='k', alpha=0.7)

            degree_rot = self.rotation_entry.get()           
            msg5 = f' rotated {degree_rot} deg'
            ax5.set_title(f'{id} {msg5}')
            fig4.tight_layout(pad=1)
            plot_id4 = "contour"
            plt.savefig(plot_id4, bbox_inches='tight')        
            img4 = Image.open("contour.png")
            plot_id4_resized = img4.resize((340, 370))
            plot_id4_resized.save("contour_resized.png")     
            img44 = ImageTk.PhotoImage(Image.open("contour_resized.png"))
            self.canvas_graph.create_image(0, 35, anchor="nw", image=img44)
            os.remove("contour_resized.png")
            plt.clf()
            plt.close(fig4)
            
# cross_section
        degree_theta = float(self.rotation_entry.get())
        if (degree_theta == 0):
            fig3 = plt.figure(figsize=(2.3, 1.8))     
            spec3 = gridspec.GridSpec(ncols=1, nrows=1)
            ax4 = fig3.add_subplot(spec3[0, 0])
            ax4.set_xlabel('Cross section (mm)')
            ax4.set_ylabel('Film thickness (Å)')
            ax4.set(title=id)
        
            value_theta = 0
            radian_theta = math.radians(value_theta)

            x_ori = df["A"]
            y_ori = df["B"]   
            z = df["Z"]
        
            x_theta = math.cos(radian_theta)*x_ori + math.sin(radian_theta)*y_ori
            y_theta = - math.sin(radian_theta)*x_ori + math.cos(radian_theta)*y_ori
            n3 = 80
            xi_theta = np.linspace(np.min(x_theta), np.max(x_theta), n3)
            yi_theta = np.linspace(np.min(y_theta), np.max(y_theta), n3)
            X_theta, Y_theta = np.meshgrid(xi_theta, yi_theta)
            Y_theta = 0
            Z_theta = griddata((x_theta, y_theta), z, (X_theta, Y_theta), method="cubic")
            value_theta_90 = value_theta + 90
            radian_theta_90 = math.radians(value_theta_90)     
            x_theta_90 = math.cos(radian_theta_90)*x_ori + math.sin(radian_theta_90
                                )*y_ori
            y_theta_90 = - math.sin(radian_theta_90)*x_ori + math.cos(radian_theta_90
                                )*y_ori
            Y_theta = 0
            Z_theta_90 = griddata((x_theta_90, y_theta_90), z, (X_theta, Y_theta),
                                  method="cubic")
            ax4.plot(X_theta, Z_theta, 'o', color='green', markersize=0.7, alpha=0.01)
            ax4.plot(X_theta, Z_theta_90, 'o', color='orange', markersize=0.7, alpha=0.01)
            ax4.legend(("X-X: Green", "Y-Y: Orange"), frameon=False, loc='upper right')
        
            top_x_theta = max(x_theta)
            btn_x_theta = min(x_theta)
            ax4.set_xlim(btn_x_theta, top_x_theta)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_y_pre = self.limits_entry.get()
                limits_y = abs(float(limits_y_pre))
                btn_y_theta = Ave - limits_y/2
                top_y_theta = Ave + limits_y/2
                ax4.set_ylim(btn_y_theta, top_y_theta)            
            else:
                top_y_theta = Ave + Range*1.6
                btn_y_theta = Ave - Range*2   
                ax4.set_ylim(btn_y_theta, top_y_theta)

            fig3.tight_layout(pad=1)    
            plot_id3 ="cross_section"
            plt.savefig(plot_id3, bbox_inches='tight')
            img3 = Image.open("cross_section.png")
            plot_id3_resized = img3.resize((330, 250))
            plot_id3_resized.save("cross_resized.png")     
            img33 = ImageTk.PhotoImage(Image.open("cross_resized.png"))
            self.canvas_graph.create_image(710, 60, anchor="nw", image=img33)
            os.remove("cross_resized.png")
            plt.clf()
            plt.close(fig3)

        else:
            fig3r = plt.figure(figsize=(2.3, 1.8))     
            spec3r = gridspec.GridSpec(ncols=1, nrows=1)
            degree_theta_rot = self.rotation_entry.get()          
            ax4r = fig3r.add_subplot(spec3r[0, 0])
            ax4r.set_xlabel('On scanline (mm)')
            ax4r.set_ylabel('Film thickness (Å)')
            msg4r = f' rotated {degree_theta_rot} deg'
            ax4r.set_title(f'{id} {msg4r}')

            value_theta_r = float(self.rotation_entry.get())
            radian_theta_r = math.radians(value_theta_r)
            x_ori_r = df["A"]
            y_ori_r = df["B"]
            z = df["Z"]
            
            x_theta_r = math.cos(radian_theta_r)*x_ori_r + math.sin(
                                radian_theta_r)*y_ori_r
            y_theta_r = - math.sin(radian_theta_r)*x_ori_r + math.cos(
                                radian_theta_r)*y_ori_r
            n3 = 80
            xi_theta_r = np.linspace(np.min(x_theta_r), np.max(x_theta_r), n3)
            yi_theta_r = np.linspace(np.min(y_theta_r), np.max(y_theta_r), n3)
            X_theta_r, Y_theta_r = np.meshgrid(xi_theta_r, yi_theta_r)
            Y_theta_r = 0
            Z_theta_r = griddata((x_theta_r, y_theta_r), z, (X_theta_r, Y_theta_r),
                                method="cubic")

            value_oth_r = value_theta_r + 90
            radian_oth_r = math.radians(value_oth_r)
            x_oth_r = math.cos(radian_oth_r)*x_ori_r + math.sin(radian_oth_r)*y_ori_r
            y_oth_r = - math.sin(radian_oth_r)*x_ori_r + math.cos(radian_oth_r)*y_ori_r
            Y_theta_r = 0
            Z_oth_r = griddata((x_oth_r, y_oth_r), z, (X_theta_r, Y_theta_r),
                                method="cubic")
            ax4r.plot(X_theta_r, Z_theta_r, 'o', color='Green', markersize=0.7, alpha=0.01)
            ax4r.plot(X_theta_r, Z_oth_r, 'o', color='Orange', markersize=0.7, alpha=0.01)
            ax4r.legend(("X-X olive", "Y-Y orange"), frameon=False, loc='upper right')

            top_x_oth = max(x_theta_r)
            btn_x_oth = min(x_theta_r)
            ax4r.set_xlim(btn_x_oth, top_x_oth)
           
            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_y_oth_pre = self.limits_entry.get()
                limits_y_oth = abs(float(limits_y_oth_pre))
                btn_y_oth = Ave - limits_y_oth/2
                top_y_oth = Ave + limits_y_oth/2
                ax4r.set_ylim(btn_y_oth, top_y_oth)            
            else:
                top_y_oth = Ave + Range*1.6
                btn_y_oth = Ave - Range*2   
                ax4r.set_ylim(btn_y_oth, top_y_oth)

            fig3r.tight_layout(pad=1)    
            plot_id3r ="cross_section"
            plt.savefig(plot_id3r, bbox_inches='tight')
            img3r = Image.open("cross_section.png")
            plot_id3r_resized = img3r.resize((330, 250))
            plot_id3r_resized.save("cross_resized2.png")     
            img33r = ImageTk.PhotoImage(Image.open("cross_resized2.png"))
            self.canvas_graph.create_image(710, 60, anchor="nw", image=img33r)
            os.remove("cross_resized2.png")
            plt.clf()
            plt.close(fig3r)

# 3d_contour
        if (value_rotation == 0):
            fig3d = plt.figure(figsize=(2.5, 2.5))
            ax3d = plt.axes(projection="3d")
            ax3d.view_init(30, 240)

            x3d = df["A"]
            y3d = df["B"]
            z3d = df["Z"] 

            x3d_grid = np.linspace(np.min(x3d), np.max(x3d), 200)
            y3d_grid = np.linspace(np.min(y3d), np.max(y3d), 200)
            X3d, Y3d = np.meshgrid(x3d_grid, y3d_grid)
        
            Z3d = griddata((x3d, y3d), z3d, (X3d, Y3d), method="cubic")
            ax3d.contour3D(X3d, Y3d, Z3d, 700, cmap='turbo', alpha=0.7, antialiased=False)
 
            top_limitx3d = np.mean(x3d) + (max(x3d)-min(x3d))*0.6
            btn_limitx3d = np.mean(x3d) - (max(x3d)-min(x3d))*0.6
            ax3d.set_xlim(btn_limitx3d, top_limitx3d)

            top_limity3d = np.mean(y3d) + (max(y3d)-min(y3d))*0.6
            btn_limity3d = np.mean(y3d) - (max(y3d)-min(y3d))*0.6
            ax3d.set_ylim(btn_limity3d, top_limity3d)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_z_pre = self.limits_entry.get()
                limits_z = abs(float(limits_z_pre))
                btn_limitz3d = np.mean(z3d) - limits_z/2
                top_limitz3d = np.mean(z3d) + limits_z/2
                ax3d.set_zlim(btn_limitz3d, top_limitz3d)          
            else:
                top_limitz3d = np.mean(z3d) + (max(z3d)-min(z3d))*1.5
                btn_limitz3d = np.mean(z3d) - (max(z3d)-min(z3d))*4.5
                ax3d.set_zlim(btn_limitz3d, top_limitz3d)

            cset = ax3d.contour(X3d, Y3d, Z3d, 3, zdir='x', offset=-btn_limitx3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3d.contour(X3d, Y3d, Z3d, 3, zdir='y', offset=-btn_limity3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3d.contour(X3d, Y3d, Z3d, 10, zdir='z', offset=btn_limitz3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.9)
       
            ax3d.set_xlabel('X (mm)')
            ax3d.set_ylabel('Y (mm)')
            ax3d.set_zlabel('Thickness (Å)')
            ax3d.set_title(id)
            fig3d.tight_layout(pad=1)    
            plot_id3d ="3d"
            plt.savefig(plot_id3d, bbox_inches='tight')
            img3d = Image.open("3d.png")
            plot_id3d_resized = img3d.resize((350, 350))
            plot_id3d_resized.save("3d_resized.png")     
            img33d = ImageTk.PhotoImage(Image.open("3d_resized.png"))
            self.canvas_graph.create_image(355, 0, anchor="nw", image=img33d)
            os.remove("3d_resized.png")
            plt.clf()
            plt.close(fig3d)
            
        else:     
            fig3dr = plt.figure(figsize=(2.5, 2.5))
            ax3dr = plt.axes(projection="3d")
            ax3dr.view_init(30, 240)

            x3d = math.cos(radian_rotation)*x_orig + math.sin(radian_rotation)*y_orig
            y3d = - math.sin(radian_rotation)*x_orig + math.cos(radian_rotation)*y_orig
            z3d = df["Z"] 

            x3d_grid = np.linspace(np.min(x3d), np.max(x3d), 200)
            y3d_grid = np.linspace(np.min(y3d), np.max(y3d), 200)
            
            X3d, Y3d = np.meshgrid(x3d_grid, y3d_grid)
            Z3d = griddata((x3d, y3d), z3d, (X3d, Y3d), method="cubic")
            ax3dr.contour3D(X3d, Y3d, Z3d, 700, cmap='turbo', alpha=0.7,
                                antialiased=False)
 
            top_limitx3d = np.mean(x3d) + (max(x3d)-min(x3d))*0.6
            btn_limitx3d = np.mean(x3d) - (max(x3d)-min(x3d))*0.6
            ax3dr.set_xlim(btn_limitx3d, top_limitx3d)

            top_limity3d = np.mean(y3d) + (max(y3d)-min(y3d))*0.6
            btn_limity3d = np.mean(y3d) - (max(y3d)-min(y3d))*0.6
            ax3dr.set_ylim(btn_limity3d, top_limity3d)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_z_pre = self.limits_entry.get()
                limits_z = abs(float(limits_z_pre))
                btn_limitz3d = np.mean(z3d) - limits_z/2
                top_limitz3d = np.mean(z3d) + limits_z/2
                ax3dr.set_zlim(btn_limitz3d, top_limitz3d)          
            else:
                top_limitz3d = np.mean(z3d) + (max(z3d)-min(z3d))*1.5
                btn_limitz3d = np.mean(z3d) - (max(z3d)-min(z3d))*4.5
                ax3dr.set_zlim(btn_limitz3d, top_limitz3d)

            cset = ax3dr.contour(X3d, Y3d, Z3d, 3, zdir='x', offset=-btn_limitx3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3dr.contour(X3d, Y3d, Z3d, 3, zdir='y', offset=-btn_limity3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3dr.contour(X3d, Y3d, Z3d, 10, zdir='z', offset=btn_limitz3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.9)
       
            ax3dr.set_xlabel('X (mm)')
            ax3dr.set_ylabel('Y (mm)')
            ax3dr.set_zlabel('Thickness (Å)')
            msg3dr = f' rotated {degree_theta_rot} deg'
            ax3dr.set_title(f'{id} {msg3dr}')
            fig3dr.tight_layout(pad=1)    
            plot_id3dr ="3d"
            plt.savefig(plot_id3dr, bbox_inches='tight')
            img3dr = Image.open("3d.png")
            plot_id3dr_resized = img3dr.resize((350, 350))
            plot_id3dr_resized.save("3dr_resized.png")     
            img33dr = ImageTk.PhotoImage(Image.open("3dr_resized.png"))
            self.canvas_graph.create_image(355, 0, anchor="nw", image=img33dr)
            os.remove("3dr_resized.png")
            plt.clf()
            plt.close(fig3dr)




# histogram
        fig10 = plt.figure(figsize=(2.3, 2.3))
        spec10 = gridspec.GridSpec(ncols=1, nrows=1)       
        ax10 = fig10.add_subplot(spec10[0, 0])

        x10 = df["Z"]
        y10 = df["Z"]
        #bins = round(len(y10)/5)
        #plt.hist(y10, bins)
        

        ax10.plot(x10, y10, 'o', markersize=1.2, alpha=0.7)
        ax10.set_xlabel('Thickness (Å)', fontsize=5)
        ax10.set_ylabel('Counts', fontsize=5)
        ax10.set(title = 'Histogram')
        
        



        fig10.tight_layout(pad=1)
        plot_id10 ="hist"
        plt.savefig(plot_id10, bbox_inches='tight')
        img10 = Image.open("hist.png")
        plot_id10_resized = img10.resize((370, 260))
        plot_id10_resized.save("hist_resized.png")      
        img10 = ImageTk.PhotoImage(Image.open("hist_resized.png")) 
        self.canvas_graph.create_image(1140, 55, anchor="nw", image=img10)
        os.remove("hist_resized.png")
        plt.clf()
        plt.close(fig10)

            

# scatter1
        fig5 = plt.figure(figsize=(2.3, 1.7))
        spec5 = gridspec.GridSpec(ncols=1, nrows=1)

        df["radius (mm)"] = np.sqrt(df["A"]**2 + df["B"]**2)
        df["theta_pre (degree)"] = np.arctan2(df["B"], df["A"])*180/np.pi
        df["theta (degree)"] = df["theta_pre (degree)"].apply(lambda x: x + 360
                                if x < 0 else x)
        
        ax5 = fig5.add_subplot(spec5[0, 0])       
        a1 = 1
        b1 = 0.1
        df["u"] = a1*df["radius (mm)"] + b1*df["theta (degree)"]
        df = df.sort_values(by="u")
        df["Point"] = range(1, 1 + len(df["u"]))
        x1 = df["Point"] 
        y1 = df["Z"] 

        ax5.plot(x1, y1, 'o', markersize=1.2, alpha=0.7)
        ax5.set_xlabel('Measurement point', fontsize=5)
        ax5.set_ylabel('Film thickness (Å)', fontsize=5)
        ax5.set(title = 'As radius increasing')

        check_box = self.var_limits.get() 
        if (check_box == 1):
            limit1_bt_pre = self.limits_entry.get()
            limit1_bt = abs(float(limit1_bt_pre))
            btn_limit1 = Ave - limit1_bt/2
            top_limit1 = Ave + limit1_bt/2
            ax5.set_ylim(btn_limit1, top_limit1)
        else:
            top_limit1 = Ave + Range*1.6
            btn_limit1 = Ave - Range*2
            ax5.set_ylim(btn_limit1, top_limit1)

        fig5.tight_layout(pad=1)
        plot_id5 ="scatter1"
        plt.savefig(plot_id5, bbox_inches='tight')
        img5 = Image.open("scatter1.png")
        plot_id5_resized = img5.resize((370, 260))
        plot_id5_resized.save("scatter1_resized.png")      
        img5 = ImageTk.PhotoImage(Image.open("scatter1_resized.png")) 
        self.canvas_graph.create_image(1040, 55, anchor="nw", image=img5)
        os.remove("scatter1_resized.png")
        plt.clf()
        plt.close(fig5)

# scatter2
        fig6 = plt.figure(figsize=(2.3, 1.7))
        spec6 = gridspec.GridSpec(ncols=1, nrows=1)

        df["radius (mm)"] = np.sqrt(df["A"]**2 + df["B"]**2)
        df["theta_pre (degree)"] = np.arctan2(df["B"], df["A"])*180/np.pi
        df["theta (degree)"] = df["theta_pre (degree)"].apply(lambda x: x + 360
                                if x < 0 else x)

        ax6 = fig6.add_subplot(spec6[0, 0]) 
        a2 = 0.1
        b2 = 1
        df["v"] = a2*df["radius (mm)"] + b2*df["theta (degree)"]    
        df = df.sort_values(by="v")
        df["Point"] = range(1, 1 + len(df["v"]))
        x2 = df["Point"] 
        y2 = df["Z"] 

        ax6.plot(x2, y2, 'o', color='red', markersize=1.2, alpha=0.5)
        ax6.set_xlabel('Measurement point', fontsize=5)
        ax6.set_ylabel('Film thickness (Å)', fontsize=5)
        ax6.set(title='As theta increasing')

        check_box = self.var_limits.get() 
        if (check_box == 1):
            limit2_bt_pre = self.limits_entry.get()
            limit2_bt = abs(float(limit2_bt_pre))
            btn_limit2 = Ave - limit2_bt/2
            top_limit2 = Ave + limit2_bt/2
            ax6.set_ylim(btn_limit2, top_limit2)
        else:
            top_limit2 = Ave + Range*1.6
            btn_limit2 = Ave - Range*2
            ax6.set_ylim(btn_limit2, top_limit2)

        fig6.tight_layout(pad=1)
        plot_id6 ="scatter2"
        plt.savefig(plot_id6, bbox_inches='tight')
        img6 = Image.open("scatter2.png")
        plot_id6_resized = img6.resize((370, 260))
        plot_id6_resized.save("scatter2_resized.png")      
        img6 = ImageTk.PhotoImage(Image.open("scatter2_resized.png")) 
        self.canvas_graph.create_image(1410, 55, anchor="nw", image=img6)
        os.remove("scatter2_resized.png")
        plt.clf()
        plt.close(fig6)

        os.remove("scatter2.png")
        os.remove("scatter1.png")
        
        os.remove("3d.png")
        os.remove("cross_section.png")
        os.remove("contour.png")
          
        tk.mainloop()

    def save_file(self):
        self.canvas_graph.delete('all')       
        id = self.entry_id.get()
        df_data = pd.read_excel(filepath)
              
        col_size = len(df_data.columns)
        if (col_size == 3):
            df_data.columns = ['A', 'B', 'C']
        elif (col_size == 4):
            df_data.columns = ['A', 'B', 'C', 'D']
        elif (col_size == 5):
            df_data.columns = ['A', 'B', 'C', 'D', 'E']
        elif (col_size == 6):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F']
        elif (col_size == 7):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
        elif (col_size == 8):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
        elif (col_size == 9):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
        elif (col_size == 10):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
        elif (col_size == 11):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']
        elif (col_size == 12):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L']
        elif (col_size == 13):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
        elif (col_size == 14):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
        elif (col_size == 15):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O']
        elif (col_size == 16):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P']
        elif (col_size == 17):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q']
        elif (col_size == 18):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R']
        elif (col_size == 19):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S']      
        elif (col_size == 20):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T']
        elif (col_size == 21):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T', 'U']
        elif (col_size == 22):
            df_data.columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O',
                               'P', 'Q', 'R', 'S', 'T', 'U', 'V']   
        else:
            pass

        var1 = self.variable1.get()
        if (var1 == " T1=C"):
            z1 = df_data["C"]
        elif (var1 == " T1=D"):
            z1 = df_data["D"]
        elif (var1 == " T1=E"):
            z1 = df_data["E"]
        elif (var1 == " T1=F"):
            z1 = df_data["F"]
        elif (var1 == " T1=G"):
            z1 = df_data["G"]
        elif (var1 == " T1=H"):
            z1 = df_data["H"]
        elif (var1 == " T1=I"):
            z1 = df_data["I"]
        elif (var1 == " T1=J"):
            z1 = df_data["J"]
        elif (var1 == " T1=K"):
            z1 = df_data["K"]
        elif (var1 == " T1=L"):
            z1 = df_data["L"]
        elif (var1 == " T1=M"):
            z1 = df_data["M"]
        elif (var1 == " T1=N"):
            z1 = df_data["N"]
        elif (var1 == " T1=O"):
            z1 = df_data["O"]
        elif (var1 == " T1=P"):
            z1 = df_data["P"]
        elif (var1 == " T1=Q"):
            z1 = df_data["Q"]
        elif (var1 == " T1=R"):
            z1 = df_data["R"]
        elif (var1 == " T1=S"):
            z1 = df_data["S"]
        elif (var1 == " T1=T"):
            z1 = df_data["T"]
        elif (var1 == " T1=U"):
            z1 = df_data["U"]
        elif (var1 == " T1=V"):
            z1 = df_data["V"]    
        else:
            pass

        var2 = self.variable2.get()
        if (var2 == " T2=0"):
            z2 = 0
        elif (var2 == " T2=C"):
            z2 = df_data["C"]
        elif (var2 == " T2=D"):
            z2 = df_data["D"]
        elif (var2 == " T2=E"):
            z2 = df_data["E"]
        elif (var2 == " T2=F"):
            z2 = df_data["F"]
        elif (var2 == " T2=G"):
            z2 = df_data["G"]
        elif (var2 == " T2=H"):
            z2 = df_data["H"]
        elif (var2 == " T2=I"):
            z2 = df_data["I"]
        elif (var2 == " T2=J"):
            z2 = df_data["J"]
        elif (var2 == " T2=K"):
            z2 = df_data["K"]
        elif (var2 == " T2=L"):
            z2 = df_data["L"]
        elif (var2 == " T2=M"):
            z2 = df_data["E"]
        elif (var2 == " T2=N"):
            z2 = df_data["F"]
        elif (var2 == " T2=O"):
            z2 = df_data["O"]
        elif (var2 == " T2=P"):
            z2 = df_data["P"]
        elif (var2 == " T2=Q"):
            z2 = df_data["Q"]
        elif (var2 == " T2=R"):
            z2 = df_data["R"]
        elif (var2 == " T2=S"):
            z2 = df_data["S"]
        elif (var2 == " T2=T"):
            z2 = df_data["T"]
        elif (var2 == " T2=U"):
            z2 = df_data["U"]
        elif (var2 == " T2=V"):
            z2 = df_data["V"]
        else:
            pass

# data filter
        z = z1 - z2
        data = [df_data["A"], df_data["B"], z]
        headers = ["A", "B", "Z"]
        orignal_df = pd.concat(data, axis=1, keys=headers)
                                    
        value = self.radio_sigma.get()
        if (value == 1):
            sigma_var = float(self.entry_sigma.get())
            z_scores = np.abs(stats.zscore(orignal_df))
            df = orignal_df[(z_scores < sigma_var).all(axis=1)]
            N1 = str(orignal_df.shape[0] - df.shape[0])
        else:
            N2 = int(self.entry_outlier.get())
            orignal_df["dist"] = np.abs(orignal_df["Z"]  - np.mean(orignal_df["Z"]))
            sort_df = orignal_df.sort_values(by="dist", ascending=False)
            df = sort_df.iloc[N2:]
            Z = df["Z"]
            stdev = np.std(Z)
            top = max(df["dist"])
            sigma2 = round(top/stdev, 1)
            
# create folder and excel
        now = datetime.now()
        dt_string = now.strftime("%m-%d-%Y %H-%M-%S")      
        check_dir = os.path.isdir('./Saved_graphs')
        if (check_dir == FALSE):
            os.makedirs('Saved_graphs')
        else:
            pass
        save_path = './Saved_graphs'
        file_name = id + "  " + dt_string + ".xlsx"
        complete_name = os. path. join(save_path, file_name)
        workbook = xlsxwriter.Workbook(complete_name)
        worksheet = workbook.add_worksheet()
        worksheet.set_column('A:A', 12)
        worksheet.set_column('B:B', 10)

# contour
        value_rotation = float(self.rotation_entry.get())            
        if (value_rotation == 0):

            x = df["A"]
            y = df["B"]
            z = df["Z"] 

            fig1 = plt.figure(figsize=(2.5, 2.5))
            spec1 = gridspec.GridSpec(ncols=1, nrows=1)
            ax1 = fig1.add_subplot(spec1[0, 0])
            n = 700

            x_grid = np.linspace(np.min(x), np.max(x), n)
            y_grid = np.linspace(np.min(y), np.max(y), n)
            X, Y = np.meshgrid(x_grid, y_grid)
            Z = griddata((x, y), z, (X, Y), method="cubic")
            Ave = round(np.mean(z), 1)
        
            value1 = self.radio_value_1.get()
            if (value1 == 2):
                pass
            else:
                time_pre = self.entry_run_time.get()
                if (time_pre):
                    time = float(self.entry_run_time.get())
                    Rate_pre = abs(Ave)*(60/time)
                    Rate = round(Rate_pre)
                else:
                    pass
            Std = np.std(z)
            Std_percent = round(100*Std/abs(Ave), 2)
            Max = round(max(z))
            Min = round(min(z))
            Nonu = round(0.5*100*(Max-Min)/abs(Ave), 2)
            Range = round(max(z) - min(z))

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_bt_pre = self.limits_entry.get()
                limits_bt = abs(float(limits_bt_pre))
                limits_b = Ave - limits_bt/2
                limits_t = Ave + limits_bt/2
                num = int(self.entry_contour.get())
                if (num >100):
                    num = 100
                    levels = np.linspace(limits_b, limits_t, num)
                else:
                    levels = np.linspace(limits_b, limits_t, num)
                cp = plt.contourf(X, Y, Z, levels = levels, cmap=plt.cm.turbo, alpha=0.95)
                fig1.colorbar(cp, ax=ax1, shrink=0.5, orientation="vertical")

            else:
                contour = int(self.entry_contour.get()) 
                if (contour >100):
                    contour = 100
                else:
                    pass
                cp = plt.contourf(X, Y, Z, contour, cmap=plt.cm.turbo, alpha=0.95)
                cbar1 = fig1.colorbar(cp, ax=ax1, shrink=0.5, orientation="vertical")
                cbar1.ax.locator_params(nbins=6)

            if (value1 == 2):
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'NonU %)')
                plt.text(-75,-270, 'StdD %')
                plt.text(-75,-285, 'Max')
                plt.text(-75,-300, 'Min')
                plt.text(-75,-315, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Nonu)                                      
                plt.text(45,-270, Std_percent)
                plt.text(45,-285, Max)
                plt.text(45,-300, Min)
                plt.text(45,-315, Range)

                worksheet.write('A1', 'Time')
                worksheet.write('A2', 'ID')
                worksheet.write('A3', 'Delta (Å)')
                worksheet.write('A4', 'NonU %')
                worksheet.write('A5', 'StdD %')
                worksheet.write('A6', 'Max')
                worksheet.write('A7', 'Min')
                worksheet.write('A8', 'Range')
                worksheet.write('B1', dt_string)
                worksheet.write('B2', id)
                worksheet.write('B3', abs(Ave))
                worksheet.write('B4', Nonu)
                worksheet.write('B5', Std_percent)
                worksheet.write('B6', Max)
                worksheet.write('B7', Min)
                worksheet.write('B8', Range)
            else:  
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'Rate (Å/min)')
                plt.text(-75,-270, 'NonU %')
                plt.text(-75,-285, 'StdD %')
                plt.text(-75,-300, 'Max')
                plt.text(-75,-315, 'Min')
                plt.text(-75,-330, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Rate)                                      
                plt.text(45,-270, Nonu)
                plt.text(45,-285, Std_percent)
                plt.text(45,-300, Max)
                plt.text(45,-315, Min)
                plt.text(45,-330, Range)

                worksheet.write('A1', 'Time')
                worksheet.write('A2', 'ID')
                worksheet.write('A3', 'Delta (Å)')
                worksheet.write('A4', 'Rate (Å/min)')
                worksheet.write('A5', 'NonU %')
                worksheet.write('A6', 'StdD %')
                worksheet.write('A7', 'Max')
                worksheet.write('A8', 'Min')
                worksheet.write('A9', 'Range')
                worksheet.write('B1', dt_string)
                worksheet.write('B2', id)
                worksheet.write('B3', abs(Ave))
                worksheet.write('B4', Rate)
                worksheet.write('B5', Nonu)
                worksheet.write('B6', Std_percent)
                worksheet.write('B7', Max)
                worksheet.write('B8', Min)
                worksheet.write('B9', Range)     
            
            value = self.radio_sigma.get()
            if (value == 1):
                plt.text(-75,-345, 'Points removed')
                plt.text(45,-345, N1)
            else:
                plt.text(-75,-345, 'Sigma:')
                plt.text(45,-345, sigma2)
        
            circ = Circle((0, 0), 150, facecolor='None', edgecolor='black', lw=0.2,
                              alpha=0.001)
            ax1.add_patch(circ)
            ax1.set_aspect('equal', adjustable='box')

            var3 = self.variable3.get()
            if (var3 == " Value"):
                for x, y ,z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    plt.annotate(label, (x, y), textcoords="offset points", xytext=(0, -3),
                             ha='center', fontsize=3.8, alpha=0.75)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            elif (var3 == " Dot"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.3)
            elif (var3 == " None"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax1.scatter(x, y, marker='o', s=1, color='k', alpha=0.0001)
            else:
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    if (z > Ave):
                        ax1.scatter(x, y, marker='+', s=13, linewidths=0.3, color='k', alpha=0.7)
                    else:
                        ax1.scatter(x, y, marker='_', s=13, linewidths=0.3, color='k', alpha=0.7)

            ax1.set_xlabel('X (mm)')
            ax1.set_ylabel('Y (mm)')    
            ax1.set_title(id)
    
            fig1.tight_layout(pad=0.1)
            plot_id1 = "contour"
            plt.savefig(plot_id1, bbox_inches='tight')        
            img1 = Image.open("contour.png")
            plot_id1_resized = img1.resize((340, 370))
            plot_id1_resized.save("contour_resized.png")     
            img11 = ImageTk.PhotoImage(Image.open("contour_resized.png"))
            self.canvas_graph.create_image(0, 35, anchor="nw", image=img11)
            worksheet.insert_image('C2', 'contour.png', {'x_scale': 1.36, 'y_scale': 1.4})
            os.remove("contour_resized.png")
            plt.clf()
            plt.close(fig1)
            
        else:                 
            fig4 = plt.figure(figsize=(2.5, 2.5))
            spec4 = gridspec.GridSpec(ncols=1, nrows=1)
            ax5 = fig4.add_subplot(spec4[0, 0])
            radian_rotation = math.radians(value_rotation)
            
            x_orig = df["A"]
            y_orig = df["B"]        
            x = math.cos(radian_rotation)*x_orig + math.sin(radian_rotation)*y_orig
            y = - math.sin(radian_rotation)*x_orig + math.cos(radian_rotation)*y_orig
            z = df["Z"]
        
            n = 700 
            x_grid = np.linspace(np.min(x), np.max(x), n)
            y_grid = np.linspace(np.min(y), np.max(y), n)
            X, Y = np.meshgrid(x_grid, y_grid)
            Z = griddata((x, y), z, (X, Y), method="cubic")
            Ave = round(np.mean(z), 1)
                    
            value1 = self.radio_value_1.get()
            if (value1 == 2):
                pass
            else:
                time_pre = self.entry_run_time.get()
                if (time_pre):
                    time = float(self.entry_run_time.get())
                    Rate = round(abs(Ave)*(60/time))
                else:
                    pass
            
            Std = np.std(z)
            Std_percent = round(100*Std/abs(Ave), 2)
            Max = round(max(z))
            Min = round(min(z))
            Nonu = round(0.5*100*(Max-Min)/abs(Ave), 2)
            Range = round(max(z) - min(z))

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_bt_pre = self.limits_entry.get()
                limits_bt = abs(float(limits_bt_pre))
                limits_b = Ave - limits_bt/2
                limits_t = Ave + limits_bt/2
                num = int(self.entry_contour.get())
                levels = np.linspace(limits_b, limits_t, num)
                cp = plt.contourf(X, Y, Z, levels = levels, cmap=plt.cm.turbo, alpha=0.95)
                fig4.colorbar(cp, ax=ax5, shrink=0.5, orientation="vertical")

            else:
                contour = int(self.entry_contour.get())
                cp = plt.contourf(X, Y, Z, contour, cmap=plt.cm.turbo, alpha=0.95)
                cbar2 = fig4.colorbar(cp, ax=ax5, shrink=0.5, orientation="vertical")
                cbar2.ax.locator_params(nbins=6)
                
            ax5.set_xlabel('X (mm)')
            ax5.set_ylabel('Y (mm)')    

            if (value1 == 2):
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'NonU %)')
                plt.text(-75,-270, 'StdD %')
                plt.text(-75,-285, 'Max')
                plt.text(-75,-300, 'Min')
                plt.text(-75,-315, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Nonu)                                      
                plt.text(45,-270, Std_percent)
                plt.text(45,-285, Max)
                plt.text(45,-300, Min)
                plt.text(45,-315, Range)

                worksheet.write('A1', 'Time')
                worksheet.write('A2', 'ID')
                worksheet.write('A3', 'Delta (Å)')
                worksheet.write('A4', 'NonU %')
                worksheet.write('A5', 'StdD %')
                worksheet.write('A6', 'Max')
                worksheet.write('A7', 'Min')
                worksheet.write('A8', 'Range')
                worksheet.write('B1', dt_string)
                worksheet.write('B2', id)
                worksheet.write('B3', abs(Ave))
                worksheet.write('B4', Nonu)
                worksheet.write('B5', Std_percent)
                worksheet.write('B6', Max)
                worksheet.write('B7', Min)
                worksheet.write('B8', Range)                
            else:  
                plt.text(-75,-240, 'Delta (Å)')
                plt.text(-75,-255, 'Rate (Å/min)')
                plt.text(-75,-270, 'NonU %')
                plt.text(-75,-285, 'StdD %')
                plt.text(-75,-300, 'Max')
                plt.text(-75,-315, 'Min')
                plt.text(-75,-330, 'Range')
                plt.text(45,-240, Ave)
                plt.text(45,-255, Rate)                                      
                plt.text(45,-270, Nonu)
                plt.text(45,-285, Std_percent)
                plt.text(45,-300, Max)
                plt.text(45,-315, Min)
                plt.text(45,-330, Range)

                worksheet.write('A1', 'Time')
                worksheet.write('A2', 'ID')
                worksheet.write('A3', 'Delta (Å)')
                worksheet.write('A4', 'Rate (Å/min)')
                worksheet.write('A5', 'NonU %')
                worksheet.write('A6', 'StdD %')
                worksheet.write('A7', 'Max')
                worksheet.write('A8', 'Min')
                worksheet.write('A9', 'Range')
                worksheet.write('B1', dt_string)
                worksheet.write('B2', id)
                worksheet.write('B3', abs(Ave))
                worksheet.write('B4', Rate)
                worksheet.write('B5', Nonu)
                worksheet.write('B6', Std_percent)
                worksheet.write('B7', Max)
                worksheet.write('B8', Min)
                worksheet.write('B9', Range)
                
            value = self.radio_sigma.get()
            if (value == 1):
                plt.text(-75,-345, 'Points removed')
                plt.text(45,-345, N1)
            else:
                plt.text(-75,-345, 'Sigma:')
                plt.text(45,-345, sigma2)
                        
            circ = Circle((0, 0), 150, facecolor='None', edgecolor='black', lw=0.2,
                          alpha=0.0001)
            ax5.add_patch(circ)
            ax5.set_aspect('equal', adjustable='box')

            var3 = self.variable3.get()
            if (var3 == " Value"):
                for x, y ,z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    plt.annotate(label, (x, y), textcoords="offset points", xytext=(0, -3),
                             ha='center', fontsize=3.8, alpha=0.75)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.001)
            elif (var3 == " Dot"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.3)
            elif (var3 == " None"):
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    ax5.scatter(x, y, marker='o', s=1, color='k', alpha=0.001)
            else:
                for x, y, z in zip(x, y, z):
                    label = "{:.0f}".format(z)
                    if (z > Ave):
                        ax5.scatter(x, y, marker='+', s=13, linewidths=0.3, color='k', alpha=0.7)
                    else:
                        ax5.scatter(x, y, marker='_', s=13, linewidths=0.3, color='k', alpha=0.7)

            degree_rot = self.rotation_entry.get()           
            msg5 = f' rotated {degree_rot} deg'
            ax5.set_title(f'{id} {msg5}')
            fig4.tight_layout(pad=1)
            plot_id4 = "contour"
            plt.savefig(plot_id4, bbox_inches='tight')        
            img4 = Image.open("contour.png")
            plot_id4_resized = img4.resize((340, 370))
            plot_id4_resized.save("contour_resized.png")     
            img44 = ImageTk.PhotoImage(Image.open("contour_resized.png"))
            self.canvas_graph.create_image(0, 35, anchor="nw", image=img44)
            worksheet.insert_image('C2', 'contour.png', {'x_scale': 1.36, 'y_scale': 1.4})
            os.remove("contour_resized.png")
            plt.clf()
            plt.close(fig4)

# cross_section
        degree_theta = float(self.rotation_entry.get())
        if (degree_theta == 0):
            fig3 = plt.figure(figsize=(2.3, 1.8))     
            spec3 = gridspec.GridSpec(ncols=1, nrows=1)
            ax4 = fig3.add_subplot(spec3[0, 0])
            ax4.set_xlabel('Cross section (mm)')
            ax4.set_ylabel('Film thickness (Å)')
            ax4.set(title=id)
        
            value_theta = 0
            radian_theta = math.radians(value_theta)

            x_ori = df["A"]
            y_ori = df["B"]   
            z = df["Z"]
        
            x_theta = math.cos(radian_theta)*x_ori + math.sin(radian_theta)*y_ori
            y_theta = - math.sin(radian_theta)*x_ori + math.cos(radian_theta)*y_ori
            n3 = 80
            xi_theta = np.linspace(np.min(x_theta), np.max(x_theta), n3)
            yi_theta = np.linspace(np.min(y_theta), np.max(y_theta), n3)
            X_theta, Y_theta = np.meshgrid(xi_theta, yi_theta)
            Y_theta = 0
            Z_theta = griddata((x_theta, y_theta), z, (X_theta, Y_theta), method="cubic")
            value_theta_90 = value_theta + 90
            radian_theta_90 = math.radians(value_theta_90)     
            x_theta_90 = math.cos(radian_theta_90)*x_ori + math.sin(radian_theta_90
                                )*y_ori
            y_theta_90 = - math.sin(radian_theta_90)*x_ori + math.cos(radian_theta_90
                                )*y_ori
            Y_theta = 0
            Z_theta_90 = griddata((x_theta_90, y_theta_90), z, (X_theta, Y_theta),
                                  method="cubic")
            ax4.plot(X_theta, Z_theta, 'o', color='green', markersize=0.7, alpha=0.01)
            ax4.plot(X_theta, Z_theta_90, 'o', color='orange', markersize=0.7, alpha=0.01)
            ax4.legend(("X-X: Green", "Y-Y: Orange"), frameon=False, loc='upper right')
        
            top_x_theta = max(x_theta)
            btn_x_theta = min(x_theta)
            ax4.set_xlim(btn_x_theta, top_x_theta)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_y_pre = self.limits_entry.get()
                limits_y = abs(float(limits_y_pre))
                btn_y_theta = Ave - limits_y/2
                top_y_theta = Ave + limits_y/2
                ax4.set_ylim(btn_y_theta, top_y_theta)            
            else:
                top_y_theta = Ave + Range*1.6
                btn_y_theta = Ave - Range*2   
                ax4.set_ylim(btn_y_theta, top_y_theta)

            fig3.tight_layout(pad=1)    
            plot_id3 ="cross_section"
            plt.savefig(plot_id3, bbox_inches='tight')
            img3 = Image.open("cross_section.png")
            plot_id3_resized = img3.resize((330, 250))
            plot_id3_resized.save("cross_resized.png")     
            img33 = ImageTk.PhotoImage(Image.open("cross_resized.png"))
            self.canvas_graph.create_image(710, 60, anchor="nw", image=img33)
            worksheet.insert_image('O3', 'cross_section.png', {'x_scale': 1.5, 'y_scale': 1.5})
            os.remove("cross_resized.png")
            plt.clf()
            plt.close(fig3)

        else:
            fig3r = plt.figure(figsize=(2.3, 1.8))     
            spec3r = gridspec.GridSpec(ncols=1, nrows=1)
            degree_theta_rot = self.rotation_entry.get()          
            ax4r = fig3r.add_subplot(spec3r[0, 0])
            ax4r.set_xlabel('On scanline (mm)')
            ax4r.set_ylabel('Film thickness (Å)')
            msg4r = f' rotated {degree_theta_rot} deg'
            ax4r.set_title(f'{id} {msg4r}')

            value_theta_r = float(self.rotation_entry.get())
            radian_theta_r = math.radians(value_theta_r)
            x_ori_r = df["A"]
            y_ori_r = df["B"]
            z = df["Z"]
            
            x_theta_r = math.cos(radian_theta_r)*x_ori_r + math.sin(
                                radian_theta_r)*y_ori_r
            y_theta_r = - math.sin(radian_theta_r)*x_ori_r + math.cos(
                                radian_theta_r)*y_ori_r
            n3 = 80
            xi_theta_r = np.linspace(np.min(x_theta_r), np.max(x_theta_r), n3)
            yi_theta_r = np.linspace(np.min(y_theta_r), np.max(y_theta_r), n3)
            X_theta_r, Y_theta_r = np.meshgrid(xi_theta_r, yi_theta_r)
            Y_theta_r = 0
            Z_theta_r = griddata((x_theta_r, y_theta_r), z, (X_theta_r, Y_theta_r),
                                method="cubic")

            value_oth_r = value_theta_r + 90
            radian_oth_r = math.radians(value_oth_r)
            x_oth_r = math.cos(radian_oth_r)*x_ori_r + math.sin(radian_oth_r)*y_ori_r
            y_oth_r = - math.sin(radian_oth_r)*x_ori_r + math.cos(radian_oth_r)*y_ori_r
            Y_theta_r = 0
            Z_oth_r = griddata((x_oth_r, y_oth_r), z, (X_theta_r, Y_theta_r),
                                method="cubic")
            ax4r.plot(X_theta_r, Z_theta_r, 'o', color='Green', markersize=0.7, alpha=0.01)
            ax4r.plot(X_theta_r, Z_oth_r, 'o', color='Orange', markersize=0.7, alpha=0.01)
            ax4r.legend(("X-X olive", "Y-Y orange"), frameon=False, loc='upper right')

            top_x_oth = max(x_theta_r)
            btn_x_oth = min(x_theta_r)
            ax4r.set_xlim(btn_x_oth, top_x_oth)
           
            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_y_oth_pre = self.limits_entry.get()
                limits_y_oth = abs(float(limits_y_oth_pre))
                btn_y_oth = Ave - limits_y_oth/2
                top_y_oth = Ave + limits_y_oth/2
                ax4r.set_ylim(btn_y_oth, top_y_oth)            
            else:
                top_y_oth = Ave + Range*1.6
                btn_y_oth = Ave - Range*2   
                ax4r.set_ylim(btn_y_oth, top_y_oth)

            fig3r.tight_layout(pad=1)    
            plot_id3r ="cross_section"
            plt.savefig(plot_id3r, bbox_inches='tight')
            img3r = Image.open("cross_section.png")
            plot_id3r_resized = img3r.resize((330, 250))
            plot_id3r_resized.save("cross_resized.png")     
            img33r = ImageTk.PhotoImage(Image.open("cross_resized.png"))
            self.canvas_graph.create_image(710, 60, anchor="nw", image=img33r)
            worksheet.insert_image('O3', 'cross_section.png', {'x_scale': 1.5, 'y_scale': 1.5})
            os.remove("cross_resized.png")
            plt.clf()
            plt.close(fig3r)

# 3d_contour
        if (value_rotation == 0):
            fig3d = plt.figure(figsize=(2.5, 2.5))
            ax3d = plt.axes(projection="3d")
            ax3d.view_init(30, 240)

            x3d = df["A"]
            y3d = df["B"]
            z3d = df["Z"] 

            x3d_grid = np.linspace(np.min(x3d), np.max(x3d), 200)
            y3d_grid = np.linspace(np.min(y3d), np.max(y3d), 200)
            X3d, Y3d = np.meshgrid(x3d_grid, y3d_grid)
        
            Z3d = griddata((x3d, y3d), z3d, (X3d, Y3d), method="cubic")
            ax3d.contour3D(X3d, Y3d, Z3d, 700, cmap='turbo', alpha=0.7, antialiased=False)
 
            top_limitx3d = np.mean(x3d) + (max(x3d)-min(x3d))*0.6
            btn_limitx3d = np.mean(x3d) - (max(x3d)-min(x3d))*0.6
            ax3d.set_xlim(btn_limitx3d, top_limitx3d)

            top_limity3d = np.mean(y3d) + (max(y3d)-min(y3d))*0.6
            btn_limity3d = np.mean(y3d) - (max(y3d)-min(y3d))*0.6
            ax3d.set_ylim(btn_limity3d, top_limity3d)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_z_pre = self.limits_entry.get()
                limits_z = abs(float(limits_z_pre))
                btn_limitz3d = np.mean(z3d) - limits_z/2
                top_limitz3d = np.mean(z3d) + limits_z/2
                ax3d.set_zlim(btn_limitz3d, top_limitz3d)          
            else:
                top_limitz3d = np.mean(z3d) + (max(z3d)-min(z3d))*1.5
                btn_limitz3d = np.mean(z3d) - (max(z3d)-min(z3d))*4.5
                ax3d.set_zlim(btn_limitz3d, top_limitz3d)

            cset = ax3d.contour(X3d, Y3d, Z3d, 3, zdir='x', offset=-btn_limitx3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3d.contour(X3d, Y3d, Z3d, 3, zdir='y', offset=-btn_limity3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3d.contour(X3d, Y3d, Z3d, 10, zdir='z', offset=btn_limitz3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.9)
       
            ax3d.set_xlabel('X (mm)')
            ax3d.set_ylabel('Y (mm)')
            ax3d.set_zlabel('Thickness (Å)')
            ax3d.set_title(id)
            fig3d.tight_layout(pad=1)    
            plot_id3d ="3d"
            plt.savefig(plot_id3d, bbox_inches='tight')
            img3d = Image.open("3d.png")
            plot_id3d_resized = img3d.resize((350, 350))
            plot_id3d_resized.save("3d_resized.png")     
            img33d = ImageTk.PhotoImage(Image.open("3d_resized.png"))
            worksheet.insert_image('I1', '3d.png', {'x_scale': 1.3, 'y_scale': 1.35})
            self.canvas_graph.create_image(355, 0, anchor="nw", image=img33d)
            os.remove("3d_resized.png")
            plt.clf()
            plt.close(fig3d)
            
        else:     
            fig3dr = plt.figure(figsize=(2.5, 2.5))
            ax3dr = plt.axes(projection="3d")
            ax3dr.view_init(30, 240)

            x3d = math.cos(radian_rotation)*x_orig + math.sin(radian_rotation)*y_orig
            y3d = - math.sin(radian_rotation)*x_orig + math.cos(radian_rotation)*y_orig
            z3d = df["Z"] 

            x3d_grid = np.linspace(np.min(x3d), np.max(x3d), 200)
            y3d_grid = np.linspace(np.min(y3d), np.max(y3d), 200)
            
            X3d, Y3d = np.meshgrid(x3d_grid, y3d_grid)
            Z3d = griddata((x3d, y3d), z3d, (X3d, Y3d), method="cubic")
            ax3dr.contour3D(X3d, Y3d, Z3d, 700, cmap='turbo', alpha=0.7,
                                antialiased=False)
 
            top_limitx3d = np.mean(x3d) + (max(x3d)-min(x3d))*0.6
            btn_limitx3d = np.mean(x3d) - (max(x3d)-min(x3d))*0.6
            ax3dr.set_xlim(btn_limitx3d, top_limitx3d)

            top_limity3d = np.mean(y3d) + (max(y3d)-min(y3d))*0.6
            btn_limity3d = np.mean(y3d) - (max(y3d)-min(y3d))*0.6
            ax3dr.set_ylim(btn_limity3d, top_limity3d)

            check_box = self.var_limits.get() 
            if (check_box == 1):
                limits_z_pre = self.limits_entry.get()
                limits_z = abs(float(limits_z_pre))
                btn_limitz3d = np.mean(z3d) - limits_z/2
                top_limitz3d = np.mean(z3d) + limits_z/2
                ax3dr.set_zlim(btn_limitz3d, top_limitz3d)          
            else:
                top_limitz3d = np.mean(z3d) + (max(z3d)-min(z3d))*1.5
                btn_limitz3d = np.mean(z3d) - (max(z3d)-min(z3d))*4.5
                ax3dr.set_zlim(btn_limitz3d, top_limitz3d)

            cset = ax3dr.contour(X3d, Y3d, Z3d, 3, zdir='x', offset=-btn_limitx3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3dr.contour(X3d, Y3d, Z3d, 3, zdir='y', offset=-btn_limity3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.5)
            cset = ax3dr.contour(X3d, Y3d, Z3d, 10, zdir='z', offset=btn_limitz3d,
                                linewidths=1, cmap=cm.turbo, alpha=0.9)
       
            ax3dr.set_xlabel('X (mm)')
            ax3dr.set_ylabel('Y (mm)')
            ax3dr.set_zlabel('Thickness (Å)')
            msg3dr = f' rotated {degree_theta_rot} deg'
            ax3dr.set_title(f'{id} {msg3dr}')
            fig3dr.tight_layout(pad=1)    
            plot_id3dr ="3d"
            plt.savefig(plot_id3dr, bbox_inches='tight')
            img3dr = Image.open("3d.png")
            plot_id3dr_resized = img3dr.resize((350, 350))
            plot_id3dr_resized.save("3d_resized.png")     
            img33dr = ImageTk.PhotoImage(Image.open("3d_resized.png"))
            self.canvas_graph.create_image(355, 0, anchor="nw", image=img33dr)
            worksheet.insert_image('I1', '3d.png', {'x_scale': 1.3, 'y_scale': 1.35})
            os.remove("3d_resized.png")
            plt.clf()
            plt.close(fig3dr)

# histogram








# scatter1
        fig5 = plt.figure(figsize=(2.3, 1.7))
        spec5 = gridspec.GridSpec(ncols=1, nrows=1)

        df["radius (mm)"] = np.sqrt(df["A"]**2 + df["B"]**2)
        df["theta_pre (degree)"] = np.arctan2(df["B"], df["A"])*180/np.pi
        df["theta (degree)"] = df["theta_pre (degree)"].apply(lambda x: x + 360
                                if x < 0 else x)
        
        ax5 = fig5.add_subplot(spec5[0, 0])       
        a1 = 1
        b1 = 0.1
        df["u"] = a1*df["radius (mm)"] + b1*df["theta (degree)"]
        df = df.sort_values(by="u")
        df["Point"] = range(1, 1 + len(df["u"]))
        x1 = df["Point"] 
        y1 = df["Z"] 

        ax5.plot(x1, y1, 'o', markersize=1.2, alpha=0.7)
        ax5.set_xlabel('Measurement point', fontsize=5)
        ax5.set_ylabel('Film thickness (Å)', fontsize=5)
        ax5.set(title = 'As radius increasing')

        check_box = self.var_limits.get() 
        if (check_box == 1):
            limit1_bt_pre = self.limits_entry.get()
            limit1_bt = abs(float(limit1_bt_pre))
            btn_limit1 = Ave - limit1_bt/2
            top_limit1 = Ave + limit1_bt/2
            ax5.set_ylim(btn_limit1, top_limit1)
        else:
            top_limit1 = Ave + Range*1.6
            btn_limit1 = Ave - Range*2
            ax5.set_ylim(btn_limit1, top_limit1)

        fig5.tight_layout(pad=1)
        plot_id5 ="scatter1"
        plt.savefig(plot_id5, bbox_inches='tight')
        img5 = Image.open("scatter1.png")
        plot_id5_resized = img5.resize((370, 260))
        plot_id5_resized.save("scatter1_resized.png")      
        img5 = ImageTk.PhotoImage(Image.open("scatter1_resized.png")) 
        self.canvas_graph.create_image(1040, 55, anchor="nw", image=img5)
        worksheet.insert_image('U3', 'scatter1.png', {'x_scale': 1.6, 'y_scale': 1.6})
        os.remove("scatter1_resized.png")
        plt.clf()
        plt.close(fig5)

# scatter2
        fig6 = plt.figure(figsize=(2.3, 1.7))
        spec6 = gridspec.GridSpec(ncols=1, nrows=1)

        df["radius (mm)"] = np.sqrt(df["A"]**2 + df["B"]**2)
        df["theta_pre (degree)"] = np.arctan2(df["B"], df["A"])*180/np.pi
        df["theta (degree)"] = df["theta_pre (degree)"].apply(lambda x: x + 360
                                if x < 0 else x)

        ax6 = fig6.add_subplot(spec6[0, 0]) 
        a2 = 0.1
        b2 = 1
        df["v"] = a2*df["radius (mm)"] + b2*df["theta (degree)"]    
        df = df.sort_values(by="v")
        df["Point"] = range(1, 1 + len(df["v"]))
        x2 = df["Point"] 
        y2 = df["Z"] 

        ax6.plot(x2, y2, 'o', color='red', markersize=1.2, alpha=0.5)
        ax6.set_xlabel('Measurement point', fontsize=5)
        ax6.set_ylabel('Film thickness (Å)', fontsize=5)
        ax6.set(title='As theta increasing')

        check_box = self.var_limits.get() 
        if (check_box == 1):
            limit2_bt_pre = self.limits_entry.get()
            limit2_bt = abs(float(limit2_bt_pre))
            btn_limit2 = Ave - limit2_bt/2
            top_limit2 = Ave + limit2_bt/2
            ax6.set_ylim(btn_limit2, top_limit2)
        else:
            top_limit2 = Ave + Range*1.6
            btn_limit2 = Ave - Range*2
            ax6.set_ylim(btn_limit2, top_limit2)

        fig6.tight_layout(pad=1)
        plot_id6 ="scatter2"
        plt.savefig(plot_id6, bbox_inches='tight')
        img6 = Image.open("scatter2.png")
        plot_id6_resized = img6.resize((370, 260))
        plot_id6_resized.save("scatter2_resized.png")      
        img6 = ImageTk.PhotoImage(Image.open("scatter2_resized.png")) 
        self.canvas_graph.create_image(1410, 55, anchor="nw", image=img6)
        worksheet.insert_image('AA3', 'scatter2.png', {'x_scale': 1.6, 'y_scale': 1.6})
        os.remove("scatter2_resized.png")
        plt.clf()
        plt.close(fig6)

        workbook.close()

        os.remove("scatter2.png")
        os.remove("scatter1.png")
        os.remove("3d.png")
        os.remove("cross_section.png")
        os.remove("contour.png")
          
        tk.mainloop()
        
def main():
    my_gui = app_gui()
main()       

# li.hou2009@gmail.com # in memory of dear father Chongjian Hou (1928-2021) #
