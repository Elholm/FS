import os
import time
from glob import glob

from tqdm import tqdm

import pandas as pd

from test_dict import fs_ejd

SLA_path_1 = r"F:\DEAS\FS\02 Tilbud\Vundne\*"
SLA_path_2 = r"F:\DEAS\FS\02 Tilbud\Tilbud\*"
print(__file__)
# dirs = glob(SLA_path_1)
accepted_answers = ["1","2"]
while True:
    choice = input("1 for Vundne, 2 for Tilbud. Write exit to exit:\n")
    # print(f"Your choice was {choice} and type {type(choice)}")
    if choice == "1":
        dirs = glob(SLA_path_1)
        break
    elif choice == "2":
        dirs = glob(SLA_path_2)
        break
    elif choice.lower() == "exit":
        exit()
    else:
        print("Try again")#, write 1 for Vundne and 2 for Tilbud")


actual_dirs = [x for x in dirs if os.path.isdir(x)]

### Construct pandas dataframe
df = pd.DataFrame()

### Blacklisted ejendomme
blacklist = ["008-222"]

job_list = ["blue","grey","green","sc","cts","feje"]
data_dict = {}
for x in job_list:
    data_dict[f"{x}_list"] = []

names = []
numbers = []

failed_dirs = []
count = 0
end_count = 1000
for x in tqdm(actual_dirs):
    if count > end_count:
        print(f"You have reached the end_count: {end_count}")
        break
    # print(x)
    files = glob(x+r"\*.xls*")
    sla_files = [y for y in files if "SLA" in y and "Kalk" not in y]
    number = "Not found"
    # print(sla_files)
    for y in job_list:
        data_dict[y] = 0
    if sla_files == []:
        # print(f"No files were found on {os.path.basename(x)}")
        failed_dirs.append(x)
    else:
        times = [os.path.getmtime(x) for x in sla_files]
        newest = times.index(max(times))
        wb = fs_ejd(sla_files[newest])
        try:
            data_dict["blue"] = wb.get_blå_timer()
        except:
            data_dict["blue"] = -1
        try:
            data_dict["grey"] = wb.get_grå_timer()
        except:
            data_dict["grey"] = -1
        try:
            data_dict["green"] = wb.get_grøn_timer()
        except:
            data_dict["green"] = -1
        try: 
            data_dict["sc"] = wb.get_SC_money()
        except:
            data_dict["sc"] = -1
        try: 
            data_dict["cts"] = wb.get_CTS_money()
        except:
            data_dict["cts"] = -1
        try: 
            data_dict["feje"] = wb.get_feje_timer()
        except:
            data_dict["feje"] = -1
        try:
            number = wb.nr
        except:
            pass
        try:
            wb.wb.close()
        except:
            pass
    for y in job_list:
        data_dict[f"{y}_list"].append(data_dict[y])
    names.append(os.path.basename(x))

    numbers.append(number)
    
    # if "xxx" in x:
        # print(f"The the folder {x.split[-1]}contains 'xxx'")
    count += 1
print(f"The amount of directories is {len(dirs)}")
print(f"The amount of actual directories is {len(actual_dirs)}")
print(f"The amount of failed dirs is {len(failed_dirs)}")
#   print(len(blue_list))
df["navne"] = names
df["ejd nr"] = numbers
for x in job_list:
    df[x] = data_dict[f"{x}_list"]
# df.to_csv("test_fs_data.csv", encoding = 'utf-8-sig')
for x in range(10):
    if os.path.exists(f"results\\test_fs_data_{x}.xls"):
        pass
    else:
        df.to_excel(f"results\\test_fs_data_{x}.xls")
        print(f"File saved as: test_fs_data_{x}.xls")
        break
df_failed = pd.DataFrame()
df_failed["names"] = failed_dirs
df_failed.to_excel("failed_dirs.xls")
