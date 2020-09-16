import openpyxl
import pandas as pd
from tqdm import tqdm
import os

# søgeord = ["moms","Faldstamme","Varmecentral","Måler","Forgæves","Boksventilator","Varmeventil","Gasmåler","Vandmåler","Varmemåler","Kedler",
# "Varmtvandsbeholder","Tæring","Påfyldning","Alarm","Fjernvarme","Pumpe","Aflæsning","Motorventil","Vaskemaskine","Tørretumbler","Vagtudkald"]
# wb = openpyxl.load_workbook("13-254 PKA.xlsx")

søgeord = ["moms","Genmontering","Levering","Punkteret","Skydebeslag","Isolering",
"Nøgler","Nedløbsrør","Flyttecylinder","Fremstilling","Forankring","Isolering",
"Dørskinne","Momskorrektion","Lift","Elevator","Maling"]
wb = openpyxl.load_workbook("13-234 PKA.xlsx")


ws = wb.active
df = pd.DataFrame(ws.values)
rm_list = []
none_nr = 0
autosum = 0
for i in df.itertuples():
    rem = False
    for x in søgeord:
        if i[6] != None:
            if x.lower() in f"{i[6]} d".lower():
                rem = True
            else:
                pass
        else:
            rem = True
            none_nr += 1
    if rem:
        rm_list.append(i[0])
        # print(f"Row {i[0]} will be deleted")
        # print(f"{i[6]}")
        # print(i)
    elif isinstance(i[11],int) or isinstance(i[11],float):
        autosum += i[11]
    else:
        pass
    if i[4] == "BilagsNr":
        # print(df.iloc[i[0],7])
        df.at[i[0],9] = autosum
        autosum = 0
    else:
        pass
        
none_nr = none_nr / len(søgeord)
print(f"{len(rm_list)} rows will be deleted, \n{none_nr:.0f} were None")

new_rm_list = [df.index[x] for x in rm_list]
df.drop(new_rm_list,inplace = True)

for x in range(10):
    if os.path.exists(f"F:/DEAS/FS/19 Driftssupport/Scripts/analysis_{x}.xlsx"):
        pass
    else:
        df.to_excel(f"F:/DEAS/FS/19 Driftssupport/Scripts/analysis_{x}.xlsx")
        print(f"File saves as: analysis_{x}.xlsx")
        break