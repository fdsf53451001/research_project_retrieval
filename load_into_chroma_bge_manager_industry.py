import json
import tqdm
import chromadb
from utils.cal_embedding_bge_zh import calculate_docs_embedding_zh
from utils.load_source_excel import get_industry_coop_proj

client = chromadb.PersistentClient(path="chroma_project_v3")
collection = client.get_or_create_collection("chroma_project_v3")

manager_group = {}

df_list = get_industry_coop_proj()
for key in df_list:
    year_data = df_list[key]

    for i in tqdm.tqdm(range(len(year_data)),desc=key):
        manager = str(year_data.iloc[i]['計畫主持人'])
        project_name = str(year_data.iloc[i]['計畫中文名稱'])
        abstract = str(year_data.iloc[i]['計畫中文摘要'])
        keywords = str(year_data.iloc[i]['中文關鍵字'])

        text = project_name + " " + abstract + " " + keywords + "\n"

        if manager not in manager_group:
            manager_group[manager] = text
        else:
            manager_group[manager] += text

# save to chroma
for manager in manager_group:
    embeddings = None
    for _ in range(3): # max retry = 3
        embeddings = calculate_docs_embedding_zh([manager_group[manager]])
        if embeddings:
            break

    collection.add(
        documents=[manager_group[manager]],
        ids=[manager],
        embeddings=embeddings,
        metadatas=[{'manager':manager}]
    )

with open('data/industry_coop/manager_group_industry.json', 'w', encoding='utf-8') as f:
    f.write(json.dumps(manager_group, ensure_ascii=False))