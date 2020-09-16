#Til LPN
import openpyxl
import tqdm
import os
#søgeord = ["faldstamme","varmecentral","måler","forgæves","boksventilator","varmeventil","gasmåler","vandmåler","varmemåler",]
søgeord = ["Faldstamme","Varmecentral","Måler","Forgæves","Boksventilator","Varmeventil","Gasmåler","Vandmåler","Varmemåler","Kedler",
"Varmtvandsbeholder","Tæring","Påfyldning","Alarm","Fjernvarme","Pumpe","Aflæsning","Motorventil","Vaskemaskine","Tørretumbler","Vagtudkald"]


wb = openpyxl.load_workbook("13-254 PKA.xlsx")
ws = wb.active

rm_list = []

for i in ws.iter_rows():
    rem = False
    for x in søgeord:
        if i[5].value != None:
            if x.lower() in f"{i[5].value} d".lower():
                rem = True
            else: 
                pass
        else:
            rem = True
        
    if rem:
        rm_list.append(i[5].row)
        print(f"Row {i[5].row} will be deleted")
print(len(rm_list))

for x in tqdm.tqdm(reversed(rm_list)):
    ws.delete_rows(x)

for x in range(10):
    if os.path.exists(f"F:/DEAS/FS/19 Driftssupport/Scripts/new_analysis_{x}.xlsx"):
        pass
    else:
        wb.save(f"F:/DEAS/FS/19 Driftssupport/Scripts/new_analysis_{x}.xlsx")
        print(f"File saved as: new_analysis_{x}.xlsx")
        break
