from langchain.vectorstores.chroma import Chroma
import pandas as pd
import tqdm
from openpyxl.worksheet.datavalidation import DataValidation

from utils.cal_embedding_bge_zh import get_embeddings_zh
from utils.load_former_manager import get_former_manager

vectorstore = Chroma("chroma_project_v2",persist_directory='chroma_project_v2',embedding_function=get_embeddings_zh())
output_excel = 'data/推薦表統合2nd_分析_v3.xlsx'

xls = pd.ExcelFile('data/推薦表統合2nd.xlsx')
former_manager = get_former_manager()

tabs = ['31DE41','E4102','E4103','E4103-2','E4104','E4105','E4105-2','E4105-3','E4106','E4108-1','E4108-2','E4110','E41']
RECOMMAND_AMOUNT = 10
SELECT_AMOUNT = 3
SELECT_BOX_SYMBOL = ['Y','Z','AA']

with pd.ExcelWriter(output_excel) as writer:
    for tab in tabs:
        page_manager_list = []

        # define column name
        df = pd.read_excel(xls, tab)
        for i in range(RECOMMAND_AMOUNT):
            df['推薦委員'+str(i+1)] = ''
            df['相關分數'+str(i+1)] = ''
        df['前任委員占比'] = ''
        for i in range(SELECT_AMOUNT):
            df['選取委員'+str(i+1)] = ''

        # process data
        for i in tqdm.tqdm(range(len(df)),desc=tab):
            manager_list = []
            project_name = df.iloc[i]['計畫名稱']
            keywords = df.iloc[i]['中文關鍵字']
            
            documents = vectorstore.similarity_search_with_relevance_scores(project_name, k=RECOMMAND_AMOUNT)
    
            for j, (doc, score) in enumerate(documents):
                df.loc[df.index[i],'推薦委員'+str(j+1)] = doc.metadata['manager']
                manager_list.append(doc.metadata['manager'])
                df.loc[df.index[i],'相關分數'+str(j+1)] = score

            page_manager_list.append(manager_list)
            df.loc[df.index[i],'前任委員占比'] = len([x for x in manager_list if x in former_manager])/RECOMMAND_AMOUNT

        df.to_excel(writer, sheet_name=tab, index=False)

        # setup dropdown list
        workbook = writer.book
        worksheet = workbook[tab]

        for j in range(SELECT_AMOUNT):
            for i, manager_list in enumerate(page_manager_list):
                data_range = ','.join(manager_list)
                dv = DataValidation(type="list", formula1=f'"{data_range}"', allow_blank=True)
                dv.add(SELECT_BOX_SYMBOL[j]+str(i+2))
                worksheet.add_data_validation(dv)

        workbook.save(output_excel)