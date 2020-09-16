import os
import time
from glob import glob

import pandas as pd

from fs_data import *

SLA_path = r"F:\DEAS\FS\02 Tilbud\Vundne\*"
dirs = glob(SLA_path)

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

failed_dirs = []
count = 0
end_count = 1000
for x in actual_dirs:
    if count > end_count:
        print(f"You have reached the end_count: {end_count}")
        break
    # print(x)
    files = glob(x+r"\*.xls*")
    sla_files = [y for y in files if "SLA" in y and "Kalk" not in y]
    # print(sla_files)
    for y in job_list:
        data_dict[y] = 0
    if sla_files == []:
        print(f"No files were found on {os.path.basename(x)}")
        failed_dirs.append(x)
    else:
        times = [os.path.getmtime(x) for x in sla_files]
        newest = times.index(max(times))
        wb = get_wb(sla_files[newest])
        if "FS Aftale" in wb.sheetnames:
            aftale = wb["FS Aftale"]
            try: 
                data_dict["blue"] = get_blå_timer(aftale)
                data_dict["grey"] = get_grå_timer(aftale)
                data_dict["green"] = get_grøn_timer(aftale)
                data_dict["sc"] = get_SC_money(aftale)
                data_dict["cts"] = get_CTS_money(aftale)
                data_dict["feje"] = get_feje_timer(aftale)
            except:
                for y in job_list:
                    data_dict[y] = -1
                pass
        else:
            pass
        wb.close()
    for y in job_list:
        data_dict[f"{y}_list"].append(data_dict[y])
    names.append(os.path.basename(x))

    # if "xxx" in x:
        # print(f"The the folder {x.split[-1]}contains 'xxx'")
    count += 1


print(f"The amount of directories is {len(dirs)}")
print(f"The amount of actual directories is {len(actual_dirs)}")
print(f"The amount of failed dirs is {len(failed_dirs)}")
#   print(len(blue_list))
df["navne"] = names
for x in job_list:
    df[x] = data_dict[f"{x}_list"]
# df.to_csv("test_fs_data.csv", encoding = 'utf-8-sig')
for x in range(10):
    if os.path.exists(f"test_fs_data_{x}.xls"):
        pass
    else:
        df.to_excel(f"test_fs_data_{x}.xls")
        break
df_failed = pd.DataFrame()
df_failed["names"] = failed_dirs
df_failed.to_excel("failed_dirs.xls")
