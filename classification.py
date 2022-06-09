"""
将原始 excel 中的需要的列提取出来 -> code + name
"""

from cmath import nan
import pandas as pd
import numpy as np
import os

in_file_paths = os.listdir('./DATA/')


def alter_one_file(in_file_path):
    sec = pd.read_excel('./DATA/' + in_file_path, keep_default_na=False)
    dir_name = in_file_path[:-5]
    out_file_path = ["./STOCK/{0}/20{1}.xlsx".format(dir_name, str(i)) for i in range(14, 21)]

    code_res = [[] for i in range(14, 21)]
    name_res = [[] for i in range(14, 21)]

    for ids, row in sec.iterrows():
        if(row['ShortName'] == ''): continue
        for i in range(14, 21):
            if(row['Accper'] == '20{}-12-31'.format(str(i))):
                code_res[i - 14].append(row['Stkcd'])
                name_res[i - 14].append(row['ShortName'])

    if not os.path.exists("./STOCK/{}".format(dir_name)):
        os.makedirs("./STOCK/{}".format(dir_name))

    for i in range(14, 21):
        df = pd.DataFrame(list(zip(code_res[i-14], name_res[i-14])), columns = ['code','name'])
        df.to_excel(out_file_path[i-14], index=False)

if __name__ == '__main__':
    print(in_file_paths)
    for path in in_file_paths:
        print(path)
        alter_one_file(path)
