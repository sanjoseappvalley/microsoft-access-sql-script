from tkinter.filedialog import askopenfilename, askdirectory
import pyodbc
import os
import pandas as pd
import tkinter as tk

def accessdb():
    newfilename = fileEntry.get()
    dateTarget = dateEntry.get()

    accessfile = askopenfilename(filetypes= [("Access Files", "*.accdb")], title = 'MS Access Database')
    location = askdirectory(title='Select location to export result Excel file')

    conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ='+accessfile+';')
    conn.setdecoding(pyodbc.SQL_CHAR, encoding='utf8')
    conn.setdecoding(pyodbc.SQL_WCHAR, encoding='utf8')
    conn.setencoding(encoding='utf8')

    sql = f"SELECT * FROM Results WHERE [Measurement_Date-Time] LIKE '{dateTarget}%' AND NOT Batch_ID = 'CONTROL'"
    df = pd.read_sql(sql, conn)

    df.sort_values(by=['ID'], inplace=True)
 
    os.chdir(location)
    df.to_csv(newfilename+'.csv', index=False)
    print('Done')
    finishLabel = tk.Label(text='***Done! You can close this window now***', fg='red', bg='yellow')
    tkcanvas.create_window(200, 250, window=finishLabel)


# Create the GUI
windowGui = tk.Tk()

tkcanvas = tk.Canvas(windowGui, width=400, height=300)
tkcanvas.pack()

dateEntry = tk.Entry(windowGui)
tkcanvas.create_window(250, 50, window=dateEntry)

dateLabel = tk.Label(text="Date (Ex: 12/2/2021)")
tkcanvas.create_window(100, 50, window=dateLabel)

fileEntry = tk.Entry(windowGui)
tkcanvas.create_window(250, 120, window=fileEntry)

fileLabel = tk.Label(text='Name of the result file')
tkcanvas.create_window(100, 120, window=fileLabel)

button1 = tk.Button(text="Open database file", command=accessdb)
tkcanvas.create_window(200, 200, window=button1)

tkcanvas.mainloop()