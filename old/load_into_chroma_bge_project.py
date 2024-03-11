import json
import tqdm
import chromadb
from utils.cal_embedding_bge_zh import calculate_docs_embedding_zh
from utils.load_source_excel import get_proj_df

client = chromadb.PersistentClient(path="chroma_project")
collection = client.get_or_create_collection("chroma_project")

df_list = get_proj_df()
for key in df_list:
    year_data = df_list[key]

    for i in tqdm.tqdm(range(len(year_data)),desc=key):
        manager = year_data.iloc[i]['計畫主持人']
        project_name = year_data.iloc[i]['計畫中文名稱']
        abstract = year_data.iloc[i]['中文摘要']
        keywords = year_data.iloc[i]['中文關鍵字']
        if_pass = year_data.iloc[i]['通過']

        text = project_name + " " + abstract + " " + keywords

        embeddings = None
        for _ in range(3): # max retry = 3
            embeddings = calculate_docs_embedding_zh([text])
            if embeddings:
                break

        collection.add(
            documents=[text],
            ids=[project_name],
            embeddings=embeddings,
            metadatas=[{'manager':manager, 'project_name':project_name,'if_pass':if_pass}]
        )
        
        
        