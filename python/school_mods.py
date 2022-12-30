import pandas as pd
import os


def create_school_mods():

    path = "python\\inputs\\school_duties.xlsx"
    excel_files = os.listdir("python\\inputs")
    
    if path in excel_files:
    #if path in [x]]
    #print(excel_file)
        columns = ["school", "am", "pm", "duty", "m", "u", "w", "t", "f"]
        cols = [0, 1, 2, 3, 4, 5, 6, 7, 8]
        sheets = [0, 1, 2, 3, 4, 5, 6, 7]
        #garages = ['Forest Glen', 'North Park', 'Chicago', 'Kedzie', '74th', '77th', '103rd']
        dfs = []
        xl = pd.read_excel(path, dtype='str')
        for g in sheets:
            sheet = xl.parse(g, skiprows=5, usecols=cols, names=columns, header=None, na_values="")
            dfs.append(sheet)
            
        schools = pd.concat(dfs)
        schools.to_csv("python\\inputs\\school_duties.csv")
        os.remove(path)