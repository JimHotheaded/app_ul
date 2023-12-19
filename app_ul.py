import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfile
import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.path import Path
from matplotlib.patches import PathPatch
from matplotlib.collections import PatchCollection
from pandastable import Table
import pandas as pd
 
    
class Window():
    def __init__(self, master):
        self.main = tk.Frame(master, background="white")

        master.title('Union Link Read .CSV File') 

        self.rightframe = tk.Frame(self.main, background="white")
        self.rightframe.pack(side=tk.LEFT)
        self.leftframe = tk.Frame(self.main, background="white")
        self.leftframe.pack(side=tk.LEFT)
 
        self.rightframeheader = tk.Frame(self.rightframe, background="white")
        self.button1 = tk.Button(self.rightframeheader, text='Import CSV',  command=self.import_csv, width=10)
        self.button1.pack(pady = (0, 5), padx = (10, 0), side = tk.LEFT)  
 
        self.button2 = tk.Button(self.rightframeheader, text='Clear',  command=self.clear, width=10)
        self.button2.pack(padx = (10, 0), pady = (0, 5), side = tk.LEFT)  
 
        # self.button3 = tk.Button(self.rightframeheader, text='Generate Plot',  command=self.generatePlot, width=10)
        # self.button3.pack(pady = (0, 5), padx = (10, 0), side = tk.LEFT)  
   
        self.button4 = tk.Button(self.rightframeheader, text='Export to Excel File',  command=self.export_to_excel, width=20)
        self.button4.pack(padx = (10,0), pady = (0,5), side = tk.LEFT)
        self.rightframeheader.pack()

        self.tableframe = tk.Frame(self.rightframe, highlightbackground="blue", highlightthickness=5)
        self.table = Table(self.tableframe, dataframe=pd.DataFrame(), width=300, height=400)
        self.table.show()
        self.tableframe.pack()
 
        self.canvas = tk.Frame(self.leftframe)
        self.fig = Figure()
        self.ax = self.fig.add_subplot(111)
        self.graph = FigureCanvasTkAgg(self.fig, self.canvas)
        self.graph.draw()
        self.graph.get_tk_widget().pack()
        self.canvas.pack(padx=(20, 0))
 
        self.main.pack()
 
 
    def import_csv(self):
        types = [("CSV files","*.csv")]
        csv_file_path = askopenfilename(initialdir = ".", title = "Open File", filetypes=types)
        tempdf = pd.DataFrame()

        tempdf = pd.read_csv(csv_file_path,
                 sep=";",
                 encoding="utf-8",
                 low_memory=False
                 )

        tempdf.replace(',', '.', regex=True, inplace=True)

        column_value = ['Dosing station 1: Line dosing units: total throughput',
              'Dosing station 1: Total throughput',
              'Dosing station 1: totalizer',
              'Extruder 1: CPU heat sink temperature',
              'Extruder 1: CPU temperature',
              'Extruder 1: Melt pressure 1',
              'Extruder 1: Screw rotation speed',
              'Extruder 1: Screw torque',
              'Extruder 1: TCU: cooling-water temperature',
              'Extruder 1: TCout adapter zone 1',
              'Extruder 1: TCout temperature zone 2',
              'Extruder 1: TCout temperature zone 3',
              'Extruder 1: TCout temperature zone 4',
              'Extruder 1: TCout temperature zone 5',
              'Extruder 1: TCout temperature zone 6',
              'Extruder 1: TCout temperature zone 7',
              'Extruder 1: TCout temperature zone 8',
              'Extruder 1: TCout temperature zone 9',
              'Extruder 1: TCout temperature zone 10',
              'Extruder 1: Torque density',
              'Extruder 1: absolute screw torque',
              'Extruder 1: adapter temperature 1',
              'Extruder 1: adapter temperature 1 deviation',
              'Extruder 1: dosing unit 1: Net weight',
              'Extruder 1: dosing unit 1: material throughput',
              'Extruder 1: dosing unit 1: rotation speed ',
              'Extruder 1: dosing unit 2: Net weight',
              'Extruder 1: dosing unit 2: material throughput',
              'Extruder 1: dosing unit 2: rotation speed ',
              'Extruder 1: dosing unit 3: Net weight',
              'Extruder 1: dosing unit 3: material throughput',
              'Extruder 1: dosing unit 3: rotation speed ',
              'Extruder 1: dosing unit 4: Net weight',
              'Extruder 1: dosing unit 4: material throughput',
              'Extruder 1: dosing unit 4: rotation speed ',
              'Extruder 1: drive unit temperature',
              'Extruder 1: heating: Heating current 1',
              'Extruder 1: heating: Heating current 2',
              'Extruder 1: lubricating oil pump: oil temperature',
              'Extruder 1: melt temperature 1',
              'Extruder 1: melt temperature 1 deviation',
              'Extruder 1: motor current',
              'Extruder 1: motor temperature',
              'Extruder 1: screw power',
              'Extruder 1: side feeder 1: Screw rotation speed',
              'Extruder 1: side feeder 2: Screw rotation speed',
              'Extruder 1: specific energy',
              'Extruder 1: specific throughput',
              'Extruder 1: switch cabinet +S1',
              'Extruder 1: switch cabinet +S2',
              'Extruder 1: switch cabinet +S3',
              'Extruder 1: temperature zone 1',
              'Extruder 1: temperature zone 1 deviation',
              'Extruder 1: temperature zone 2',
              'Extruder 1: temperature zone 2 deviation',
              'Extruder 1: temperature zone 3',
              'Extruder 1: temperature zone 3 deviation',
              'Extruder 1: temperature zone 4',
              'Extruder 1: temperature zone 4 deviation',
              'Extruder 1: temperature zone 5',
              'Extruder 1: temperature zone 5 deviation',
              'Extruder 1: temperature zone 6',
              'Extruder 1: temperature zone 6 deviation',
              'Extruder 1: temperature zone 7',
              'Extruder 1: temperature zone 7 deviation',
              'Extruder 1: temperature zone 8',
              'Extruder 1: temperature zone 8 deviation',
              'Extruder 1: temperature zone 9',
              'Extruder 1: temperature zone 9 deviation',
              'Extruder 1: temperature zone 10',
              'Extruder 1: temperature zone 10 deviation',
              'Underwater pelletizer: Adapter temperature',
              'Underwater pelletizer: Die plate temperature',
              'Underwater pelletizer: Diverter valve temperature',
              'Underwater pelletizer: Exhaust fan speed',
              'Underwater pelletizer: Melt pressure ',
              'Underwater pelletizer: Pelletizer speed',
              'Underwater pelletizer: Pelletizer torque',                  
              'Underwater pelletizer: Process water flow rate',
              'Underwater pelletizer: Process water temperature',
              'Underwater pelletizer: Water pump speed',
              'Underwater pelletizer: Water tank temperature',
              'Underwater pelletizer: melt temperature']

        tempdf[column_value] = tempdf[column_value].astype(float)

        tempdf['Date'] = pd.to_datetime(tempdf['Date'], format='%d.%m.%Y %H:%M:%S')
             
        self.table.model.df = tempdf
        self.table.model.df.columns = self.table.model.df.columns.str.lower()
        self.table.redraw()  
        self.generatePlot()

    def clear(self):
        self.table.model.df = pd.DataFrame()
        self.table.redraw()
        self.ax.clear()
        self.graph.draw_idle()
 
    def generatePlot(self):
        self.ax.clear()
        if not(self.table.model.df.empty): 
            df = self.table.model.df.copy()
            self.ax.plot(pd.to_numeric(df["x"]), pd.to_numeric(df["y"]), color ='tab:blue', picker=True, pickradius=5)   
        self.graph.draw_idle()     
    
    def export_to_excel(self):
        export_df = self.table.model.df
        excel_file_handle = asksaveasfile(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])

        if excel_file_handle:
            excel_file_path = excel_file_handle.name
            export_df.to_excel(excel_file_path, index=False)
            excel_file_handle.close()
            tk.messagebox.showinfo("Success", f"File '{excel_file_path}' saved successfully!")

root = tk.Tk()
window = Window(root)
root.mainloop()