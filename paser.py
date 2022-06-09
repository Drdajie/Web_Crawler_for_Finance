import pdfplumber
import pandas as pd

import pdfplumber
import pandas as pd
import os

# import tqdm
from tqdm import trange

# with pdfplumber.open("000766_通化金马_2015年年度报告.pdf") as pdf:
#     page = pdf.pages[14]   # 第一页的信息
#     table = page.extract_tables()
#     for t in table:
#         # 得到的table是嵌套list类型，转化成DataFrame更加方便查看和分析
#         df = pd.DataFrame(t[1:], columns=t[0])
#         print(df)


class Solution:
    def strStr(self, haystack: str, needle: str) -> int:
        def KMP(s, p):
            """
            s为主串
            p为模式串
            如果t里有p，返回打头下标
            """
            nex = getNext(p)
            i = 0
            j = 0   # 分别是s和p的指针
            while i < len(s) and j < len(p):
                if j == -1 or s[i] == p[j]: # j==-1是由于j=next[j]产生
                    i += 1
                    j += 1
                else:
                    j = nex[j]

            if j == len(p): # j走到了末尾，说明匹配到了
                return i - j
            else:
                return -1

        def getNext(p):
            """
            p为模式串
            返回next数组，即部分匹配表
            """
            nex = [0] * (len(p) + 1)
            nex[0] = -1
            i = 0
            j = -1
            while i < len(p):
                if j == -1 or p[i] == p[j]:
                    i += 1
                    j += 1
                    nex[i] = j     # 这是最大的不同：记录next[i]
                else:
                    j = nex[j]

            return nex
        
        return KMP(haystack, needle)


def process_pages(pages):
    name = ['研发人员数量（人）', '研发人员数量占比', '研发投入金额（元）', '研发投入占营业收入比例']
    df = pd.DataFrame(columns=name)
    for page in pages:
        tables = page.extract_tables()
        for table in tables:
            for line in table:
                if line[0] in name:
                    df.loc[0, line[0]] = line[2]
    return df


def process_one_file(file_path):
    solve = Solution()
    target = "研发人员数量占比"

    with pdfplumber.open(file_path) as pdf:
        res = 0
        page_zero = pdf.pages[0]
        gap = page_zero.page_number

        name = ['研发人员数量（人）', '研发人员数量占比', '研发投入金额（元）', '研发投入占营业收入比例']
        df = pd.DataFrame(columns=name)
        for page  in pdf.pages:
            text = page.extract_text()
            sit = solve.strStr(text, target)
            if(sit == -1):   continue
            else:
                last_page = pdf.pages[page.page_number - 1 - gap]
                now_page = pdf.pages[page.page_number - gap]
                next_page = pdf.pages[page.page_number + 1 - gap]
                df = process_pages([last_page, now_page, next_page])
                return df
        
        return df

if __name__ == '__main__':
    all_fields = os.listdir('./YEAR_REPORT')

    col_name = ['代码', '名称','研发人员数量（人）', '研发人员数量占比', '研发投入金额（元）', '研发投入占营业收入比例']
    for field in all_fields:
        target_year_dir = 2015
    #    target_year = 2014
        
        res = pd.DataFrame(columns=col_name)
        files = os.listdir('./YEAR_REPORT/' + field + '/' + str(target_year_dir))
        all_num = len(files)
        for i, file in enumerate(files):
            print(i, ' ', file)

            cnr = file.split('_')
            row = res.shape[0]
            res.loc[row, '代码'], res.loc[row, '名称'] = cnr[0], cnr[1]

            df = process_one_file('./YEAR_REPORT/' + field + '/' + str(target_year_dir) + '/' + file)
            if(not df.empty):
                res.iloc[row, 2:] = df.iloc[0, :]
            
        res.to_excel('./RESULT/' + field + '.xlsx')