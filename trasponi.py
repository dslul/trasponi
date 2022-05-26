#!/usr/bin/env python

import pandas as pd
import os
from threading import Thread
import glob
try:
    import Tkinter as tk
except:
    import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter import *

folder_name = "Conv"
desinenza = "_conv.xlsx"

root = Tk()
root.title('Convertitore excel indagini finanziarie')
root.resizable(False, False)
root.geometry('300x150')

menubar = Menu(root)
root.config(menu=menubar)
about_menu = Menu(menubar)
lbl_info = ttk.Label(root, text='Pronto.')
btn_select = ttk.Button(root, text='Scegli cartella')
progressbar = ttk.Progressbar(root, orient='horizontal', mode='indeterminate')

def about():
    messagebox.showinfo('Informazioni', 'Sorgente: https://github.com/dslul/trasponi\n\n\nRealizzato da:\n\nDaniele Laudani <daniele.laudani@adm.gov.it>\nSalvatore Caligiore <salvatore.caligiore@adm.gov.it>\nGiovanni Farchione <giovanni.farchione@adm.gov.it>')

def convert():
    folder = filedialog.askdirectory()
    print('Cartella: ' + folder)
    if folder != "":
        lbl_info.config(text="Conversione in corso...")
        progressbar.start()
        btn_select.config(state=tk.DISABLED)
        t = Thread(target=transpose, args=(folder,), daemon=True)
        t.start()
        #transpose(folder)

def transpose(folder):
    for file in glob.glob(folder + '/**/*.xls*', recursive=True):
        #evita di processare eventuali file convertiti
        #print(os.path.basename(os.path.normpath(os.path.dirname(file))))
        if desinenza in file or os.path.basename(os.path.normpath(os.path.dirname(file))) == folder_name:
            continue
        
        dont_convert = False
        print("File: " + file)
        num_sheets = len(pd.read_excel(file, sheet_name=None))
        xls = pd.ExcelFile(file)
        
        #salta file senza conti correnti
        if num_sheets <= 2:
            continue
        
        # info intestatario conto
        df = pd.read_excel(xls, 0)
        nome = ''
        cognome = ''
        cf = ''
        banca = ''
        for index, row in df.iterrows():
            if 'Cognome' in str(row[0]):
                cognome = row[1]
            if 'Nome' in str(row[0]):
                nome = row[1]
            if 'Codice Fiscale' in str(row[0]):
                cf = row[1]
            if 'Operatore finanziario' in str(row[0]):
                banca = row[1]
        print(nome + ' ' + cognome + ' - ' + cf)
        nome_file = nome + ' ' + cognome + "_" + banca;

        if banca == "":
            continue

        dataframes = []
        for i in range(2, num_sheets):
            #individua inizio operazioni
            df = pd.read_excel(xls, i)
            #numero_conto = df.columns[1]

            start_from = 0
            n_rows = 0
            ops_found = False
            for index, row in df.iterrows():
                    n_rows = 0
                    if type(row[0]) == str and "Elenco operazioni" in row[0]:
                            n_rows = row[0][row[0].find("(")+1:row[0].find(")")]
                            start_from = index + 1
                            ops_found = True
                            break
            if ops_found == False:
                # in caso di file già convertito esci
                if i == 2:
                    dont_convert = True
                    break
                #ricopia fogli con movimenti extraconto
                foglio = pd.read_excel(file, sheet_name=xls.sheet_names[i])
                #cancella prima riga
                if "operazioni Extraconto" in str(foglio.columns[0]):
                    new_header = foglio.iloc[0] #grab the first row for the header
                    foglio = foglio[1:] #take the data less the header row
                    foglio.columns = new_header #set the header row as the df header
                dataframes.append((xls.sheet_names[i], foglio))
                continue
            
            table = {}
            for index, row in df.iloc[start_from:].iterrows():
                    #print(row[0])
                    if row[0] == "" or str(row[0]) == "nan":
                            break
                    if row[0] not in table:
                            table[row[0]] = []
                    table[row[0]].append(row[1])

            #assert len(table.keys()) == 11
            # in caso di nessuna operazione
            if n_rows == "0":
                table['Operazioni'] = ['Nessuna']
            
            df2 = pd.DataFrame(table)
            dataframes.append((xls.sheet_names[i], df2))

        
        if dont_convert == True:
            print("FILE GIÀ CONVERTITO")
            continue
        # crea se non esiste la cartella dei file convertiti
        result_dir = os.path.join(os.path.dirname(os.path.abspath(file)), folder_name)
        if os.path.exists(result_dir) == False:
            os.makedirs(result_dir, exist_ok=True)
        
        #salva
        foglio1 = pd.read_excel(file, sheet_name='Dati generali')
        foglio2 = pd.read_excel(file, sheet_name='Annotazioni')

        newfile = os.path.splitext(file)[0] + desinenza
        with pd.ExcelWriter(os.path.join(result_dir, os.path.basename(newfile))) as writer:
            foglio1.to_excel(writer, sheet_name='Dati generali', index=False)
            # imposta larghezza
            for i, col in enumerate(foglio1.columns):
                column_len = len(col) * 4
                writer.sheets['Dati generali'].set_column(i, i, column_len)
            foglio2.to_excel(writer, sheet_name='Annotazioni', index=False)
            
            for df in dataframes:
                df[1].to_excel(writer, sheet_name=df[0], index=False)
                #regola larghezza colonne
                worksheet = writer.sheets[df[0]]
                for i, col in enumerate(df[1].columns):
                    column_len = len(col) + 2
                    worksheet.set_column(i, i, column_len)
    
    lbl_info.config(text="Terminato.")
    progressbar.stop()
    btn_select.config(state=tk.NORMAL)


btn_select.config(command=convert)
about_menu.add_command(label='Informazioni', command=about)
menubar.add_cascade(label="?", menu=about_menu, underline=0)

btn_select.pack(expand=True)
lbl_info.pack(expand=True)
progressbar.pack(expand=True)
root.mainloop()
