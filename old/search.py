from langchain.vectorstores.chroma import Chroma
import pandas as pd
import tqdm

from utils.cal_embedding_bge_zh import get_embeddings_zh

vectorstore = Chroma("chroma_project",persist_directory='chroma_project',embedding_function=get_embeddings_zh())
output_excel = 'data/推薦表統合2nd_分析_pass.xlsx'

xls = pd.ExcelFile('data/推薦表統合2nd.xlsx')

tabs = ['31DE41','E4102','E4103','E4103-2','E4104','E4105','E4105-2','E4105-3','E4106','E4108-1','E4108-2','E4110','E41']
RECOMMAND_AMOUNT = 10
ONLY_PASS_PROJECT = True

with pd.ExcelWriter(output_excel) as writer:
    for tab in tabs:
        df = pd.read_excel(xls, tab)
        for i in range(RECOMMAND_AMOUNT):
            df['推薦委員'+str(i+1)] = ''
            df['相關計畫'+str(i+1)] = ''
            df['相關分數'+str(i+1)] = ''
            df['相關計畫通過'+str(i+1)] = ''

        for i in tqdm.tqdm(range(len(df)),desc=tab):
            project_name = df.iloc[i]['計畫名稱']
            keywords = df.iloc[i]['中文關鍵字']
            
            documents = vectorstore.similarity_search_with_relevance_scores(project_name, k=RECOMMAND_AMOUNT)
            if ONLY_PASS_PROJECT:
                documents = [(doc,score) for doc,score in documents if doc.metadata['if_pass'] == 'true']
            for j, (doc, score) in enumerate(documents):
                df.loc[df.index[i],'推薦委員'+str(j+1)] = doc.metadata['manager']
                df.loc[df.index[i],'相關計畫'+str(j+1)] = doc.metadata['project_name']
                df.loc[df.index[i],'相關分數'+str(j+1)] = score
                df.loc[df.index[i],'相關計畫通過'+str(j+1)] = doc.metadata['if_pass']
    
        df.to_excel(writer, sheet_name=tab, index=False)