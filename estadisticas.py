import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import tkinter as tk
from datetime import datetime
from tkinter import ttk
from tkinter import filedialog
import locale

# Configurar el formato de moneda (usa la configuración local predeterminada del sistema)
locale.setlocale(locale.LC_ALL, '')

def format_currency(value):
    """Format a value as currency."""
    return locale.currency(value, grouping=True)[:-3]

# *** CARGA UN EXCEL EN UN DATAFRAME
def loadExcel(file):
    try:# ACA AGREGAMOS LAS COLUMNAS SI ES PAROISSIEN
        dataFrame = pd.read_excel(file, sheet_name=2, header=None, skiprows=4, usecols='B:F, I, J, O:R', 
                    names=['Fecha', 'CubiertosDia', 'PromedioDia', 'CubiertosNoche', 'PromedioNoche', 'TotalTarjeta', 'TotalEfectivo', 'deliveryDia', 'deliveryNoche', 'totalDelivery', 'importeDelivery'],
                    dtype={'B': datetime,
                        'C': np.int32,
                        'D': np.float64,
                        'E': np.int32,
                        'F': np.float64,
                        'I': np.float64,
                        'J': np.float64,
                        'O': np.int32, 
                        'P': np.int32,
                        'Q': np.int32,
                        'R': np.float64
                        }).dropna()
    except:# ACA AGREGAMOS LAS COLUMNAS para el resto
        dataFrame = pd.read_excel(file, sheet_name=2, header=None, skiprows=4, usecols='B:F, I, J', 
                    names=['Fecha', 'CubiertosDia', 'PromedioDia', 'CubiertosNoche', 'PromedioNoche', 'TotalTarjeta', 'TotalEfectivo'],
                    dtype={'B': datetime,
                        'C': np.int32,
                        'D': np.float64,
                        'E': np.int32,
                        'F': np.float64,
                        'I': np.float64,
                        'J': np.float64,
                        }).dropna()
    return dataFrame

# *** crea el cuadro
def table_builder(shops, data):
    rootTable = tk.Tk()
    rootTable.title('Titulo del Cuadro')
    style = ttk.Style()
    style.theme_use("clam")  # Use the 'clam' theme for better styling options
    style.configure("Treeview",
                    background="#ffffff",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="#ffffff",
                    font=("Arial", 10))
    style.map("Treeview",
            background=[("selected", "#009879")],
            foreground=[("selected", "white")])
    # Configure heading style
    style.configure("Treeview.Heading",
                    background="#009879",
                    foreground="white",
                    font=("Arial", 10, "bold"))
    if len(data)==10:
        columns = ('DIA', 'NOCHE', 'TOTAL', 'VENTAS', 'EFECTIVO', 'TARJETA', '% TARJETA-VENTAS',
                    'DELIVERY DIA', 'DELIVERY NOCHE', 'TOTAL DELIVERY', 'IMPORTE DELIVERY')
        wides = (80,80,80,120,120,120,80,80,80,80,120)
    else:
        columns = ('DIA', 'NOCHE', 'TOTAL', 'VENTAS', 'EFECTIVO', 'TARJETA', '% TARJETA-VENTAS')
        wides = (80,80,80,120,120,120,80)

    # Create a frame for the treeview to add a border
    frame = tk.Frame(rootTable, bg='#f0f0f0')
    frame.pack(expand=True, fill='both', padx=10, pady=10)
    tv = ttk.Treeview(frame, columns= columns)
    tv.column("#0", width=180, minwidth=150)
    tv.heading("#0", text='LOCAL', anchor=tk.CENTER)
    # Apply striped row effect
    
    for i, column in enumerate(columns):
        tv.column(column, width=wides[i], minwidth=wides[i], anchor=tk.CENTER)
        tv.heading(column, text=column, anchor=tk.CENTER)
    rowToCompare=[]
    tv.tag_configure('oddrow', background="#f8ffb8")
    tv.tag_configure('evenrow', background="#ffffff")
    for i, shop in enumerate(shops):
        rows=[]
        rowPercent=[]
        for col in range(6): #la ultima columna esta agregada
            value = data[col][i]
            value = 1 if value == 0 else value
            if rowToCompare: #valida que haya una fila anterior para comparar
                perc = value/rowToCompare[col]-1
                difference = value - rowToCompare[col]
                if col in range(3,6): # si son importes le doy formato
                    difference = format_currency(difference)
                rowPercent.append(f"{difference} ({perc:.0%})")
            rows.append(value)
        # agrega comparacion de % tarjeta-ventas
        value = rows[5]/rows[3]
        if rowToCompare: #valida que haya una fila anterior para comparar
            rowPercent.append("")
        rows.append(f'{value:.0%}') #agrega columna de % tarjeta-ventas
        # ACA AGREGAMOS LAS COLUMNAS SI ES PAROISSIEN
        if len(data)==10:
            for col in range(6,10):
                value = data[col][i]
                value = 1 if value == 0 else value
                if rowToCompare: #valida que haya una fila anterior para comparar
                    perc = value/int(rowToCompare[col+1])-1 #por la columna agregada de %
                    difference = value - int(rowToCompare[col+1])
                    rowPercent.append(f"{difference} ({perc:.0%})")
                rows.append(value)
        rowToCompare = rows.copy()
        if rowPercent: #valida que haya una fila de porcentajes
            tv.insert("",tk.END, text="Variacion", values=rowPercent, tags=('oddrow',))
        # Formatear las columnas de ventas como moneda
        rows[3] = format_currency(rows[3])  # VENTAS
        rows[4] = format_currency(rows[4])  # EFECTIVO
        rows[5] = format_currency(rows[5])  # TARJETA
        if len(data)==10: #Paroissien
            rows[9] = format_currency(rows[9])  # total delivery
        tv.insert("", tk.END, text=shop, values=rows, tags=('evenrow',))
    tv.pack(expand=True, fill='both')
    rootTable.mainloop()


# *** ESTA FUNCION CREA EL GRAFICO
def graph_builder(shops, data):
    if len(data) == 10: # es PAROISSIEN
        labels=['Cubiertos DIA', 'Cubiertos NOCHE', 'TOTAL Cubiertos',
                 'Delivery DIA', 'Delivery NOCHE', 'Total Delivery']
        max = np.max(data[8]) + 1000 #calcula el pico del grafico
        n_series = 6
        data_columns = (0,1,2,6,7,8)
    else:    
        labels=['Cubiertos DIA', 'Cubiertos Noche', 'TOTAL Cubiertos']
        max = np.max(data[2]) + 1000 #calcula el pico del grafico
        n_series = 3
        data_columns = (0,1,2)
    n_observations=len(shops)
    x = np.arange(n_observations)  # the label locations
    fig, ax = plt.subplots(layout='constrained')
    # Determine bar widths
    width_cluster = 0.7
    width_bar = width_cluster/n_series
    for n, d in enumerate(data_columns):
        x_positions = x+(width_bar*n)-width_cluster/2
        rects = ax.bar(x_positions, data[d], width_bar, align='edge', label=labels[n])
        ax.bar_label(rects, padding=5)
    ax.set_ylabel('Cubiertos')
    ax.set_title('Comparativa')
    ax.set_xticks(x , shops)
    ax.legend(loc='upper left', ncols=3)
    ax.set_ylim(0, max)
    return

# *** CALCULA CUBIERTOS
def dishesCalculator(dataFrames):
    #estos son para PAROISSIEN
    deliveryDia = []
    deliveryNoche = []
    totalDelivery = []
    importeDelivery = []
    # estos son para todos los locales
    cubiertosDia = []
    cubiertosNoche = []
    totalCubiertos = []
    sales = []
    salesCash =[]
    salesCredit = []
    try:
        for df in dataFrames:
            # estos son para PAROISSIENS
            deliveryDia.append(int(np.sum(df['deliveryDia'].to_numpy())))
            deliveryNoche.append(int(np.sum(df['deliveryNoche'].to_numpy())))
            totalDelivery.append(int(np.sum(df['totalDelivery'].to_numpy())))
            importeDelivery.append(int(np.sum(df['importeDelivery'].to_numpy())))
            # estos son para todos los locales
            cubiertosDia.append(int(np.sum(df['CubiertosDia'].to_numpy())))
            cubiertosNoche.append(int(np.sum(df['CubiertosNoche'].to_numpy())))
            salesCash.append(int(np.sum(df['TotalEfectivo'].to_numpy())))
            salesCredit.append(np.sum(df['TotalTarjeta'].to_numpy()))
            sales.append( salesCash[-1] + salesCredit[-1] )
            totalCubiertos.append(cubiertosDia[-1] + cubiertosNoche[-1])
        return [ cubiertosDia,
                 cubiertosNoche,
                 totalCubiertos,
                 sales,
                 salesCash,
                 salesCredit,
                 deliveryDia,
                 deliveryNoche,
                 totalDelivery,
                 importeDelivery]
    except:
        for df in dataFrames:
            cubiertosDia.append(int(np.sum(df['CubiertosDia'].to_numpy())))
            cubiertosNoche.append(int(np.sum(df['CubiertosNoche'].to_numpy())))
            salesCash.append(int(np.sum(df['TotalEfectivo'].to_numpy())))
            salesCredit.append(np.sum(df['TotalTarjeta'].to_numpy()))
            sales.append( salesCash[-1] + salesCredit[-1] )
            totalCubiertos.append(cubiertosDia[-1] + cubiertosNoche[-1])
        return [ cubiertosDia,
                 cubiertosNoche,
                 totalCubiertos,
                 sales,
                 salesCash,
                 salesCredit]

# Carga los DataFrames
def loadDataFrames(fileNames):
    dataFrames = []
    dataShops = []
    for fn in fileNames:
        nameStart = fn.rfind('/') + 1 #Busca el inicio del filename
        dataShops.append(fn[nameStart:-5]) #Le quita la extension
        dataFrames.append(loadExcel(fn))
    return dataFrames, dataShops

def graph():
    dataFrames, dataShops= loadDataFrames(fileNames)
    dishes = dishesCalculator(dataFrames)
    # ***** configura el grafico
    graph_builder(dataShops, dishes)
    plt.show()

def table():
    dataFrames, dataShops = loadDataFrames(fileNames)
    dishes = dishesCalculator(dataFrames)
    # ***** configura la tabla
    table_builder(dataShops, dishes)
    plt.show()

def erase(btn):
    info = btn.grid_info() #get the row
    row = info['row']
    # Remove the file name from the list
    del fileNames[row]
    # Destroy the corresponding label and button widgets
    labels[row].destroy()
    buttons[row].destroy()
    # Remove the corresponding elements from the lists
    del labels[row]
    del buttons[row]
    # Update the layout for remaining elements
    for i in range(0, len(fileNames)):
        labels[i].grid_configure(row=i)
        buttons[i].grid_configure(row=i)

def selectFile():
    filePath = filedialog.askopenfilename(title='Seleccioná el XLSXL', filetypes=[('Excel:', '*.xlsx')])
    if filePath:
        fileNames.append(filePath)
        row = len(fileNames) - 1
        label = tk.Label(filesSelector, text=fileNames[row], bg=backFrame, foreground=fontDark)
        label.grid(row=row, column=1, padx=10, pady=10, sticky='e')
        labels.append(label)
        button = tk.Button(filesSelector, text='borrar')
        button.config(command=lambda b=button: erase(b)) #add a command to call erase function
        button.grid(row=row, column=2, padx=10, pady=10)
        buttons.append(button)
 
# Aca comenzamos con la GUI************************************
backRoot = '#f0f0f0'
backFrame ='#fff'
backButt = '#1227E6'
backTable ='#12ADE6'
backTitle ='#009879'
backData = '#5F6CE6'
backPorc = '#53D9E6'
fontClear = '#f0f0f0'
fontDark = '#111111'
borderWidth = 2
relief= "groove"
root = tk.Tk()
root.title('Comparativa de Cubiertos')
root.config(background=backRoot)
filesSelector = tk.Frame(root)
filesSelector.config(width=650, height=300, bg= backFrame)
filesSelector.grid(row=0, column=0, sticky='e', padx=10, pady=10)
fileNames = []
labels = []
buttons = []
tk.Button(filesSelector, text='Seleccioná el Excel', command=selectFile, font=(18)).grid(row=0, column=0, padx=10, pady=10)
tk.Button(filesSelector, text='Gráfico', command=lambda: graph()).grid(row=1, column=0, padx=10, pady=10)
tk.Button(filesSelector, text='Cuadro', command=lambda: table()).grid(row=2, column=0, padx=10, pady=10)
root.mainloop()


