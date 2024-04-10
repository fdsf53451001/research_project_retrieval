import pandas as pd
from typing import List

def get_proj_df() -> List[pd.DataFrame]:
    # 研究計畫
    xls = pd.ExcelFile('data/(勿對外公開資料或流傳)108-112年智慧計算學門大批專題計畫申請案件(含中英文摘要及關鍵字)1130215.xlsx')
    year = ['108', '109', '110', '111', '112']
    df_list = {}
    for y in year:
        # load all data
        df1 = pd.read_excel(xls, y)
        df1.columns = df1.iloc[0]
        df1 = df1.iloc[1:]
        # print(df1.head())

        # load accepted data
        xls2 = pd.ExcelFile('data/(密件)智慧計算學門統計1130130.xlsx')
        df2 = pd.read_excel(xls2, y+'總計畫清單')
        pass_proj_name_list = df2['計畫中文名稱'].to_list()

        pass_list = []
        for i in range(len(df1)):
            if df1.iloc[i]['計畫中文名稱'] in pass_proj_name_list:
                pass_list.append('true')
            else:
                pass_list.append('false')
        df1['通過'] = pass_list  
        df_list[y] = df1

    # print(df_list['108']['通過'].value_counts())
    return df_list

def get_industry_coop_proj():
    # 產學計劃
    xls = pd.ExcelFile('data/industry_coop/108-112產學計畫E41申請名冊.xlsx')
    year = ['專題計畫綜合查詢']
    df_list = {}
    for y in year:
        # load all data
        df1 = pd.read_excel(xls, y)
        df_list[y] = df1

    return df_list

print(len(get_industry_coop_proj()))