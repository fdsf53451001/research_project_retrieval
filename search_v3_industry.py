from langchain.vectorstores.chroma import Chroma
import pandas as pd
import tqdm
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting import Rule
from openpyxl.formatting.rule import ColorScaleRule
# from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill
from openpyxl.styles.differential import DifferentialStyle

from utils.cal_embedding_bge_zh import get_embeddings_zh
from utils.load_former_manager import get_former_manager

vectorstore = Chroma("chroma_project_v3",persist_directory='chroma_project_v3',embedding_function=get_embeddings_zh())
output_excel = 'data/industry_coop/推薦表統合2nd_分析.xlsx'

xls = pd.ExcelFile('data/industry_coop/推薦表統合2nd.xlsx')
former_manager = get_former_manager(file_path='data/industry_coop/曾任委員.txt')

tabs = ['31DE41']
RECOMMAND_AMOUNT = 10
SELECT_AMOUNT = 3

def search_and_analyze():
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

            add_recommenad_manager_dropdown_list(writer, tab, page_manager_list)
            highligh_former_manager(writer, tab)
            draw_color_for_similarity_score(writer, tab)


def add_recommenad_manager_dropdown_list(writer, tab, page_manager_list):
    # setup dropdown list
    SELECT_BOX_SYMBOL = ['Y','Z','AA']
    workbook = writer.book
    worksheet = workbook[tab]

    for j in range(SELECT_AMOUNT):
        for i, manager_list in enumerate(page_manager_list):
            data_range = ','.join(manager_list)
            dv = DataValidation(type="list", formula1=f'"{data_range}"', allow_blank=True)
            dv.add(SELECT_BOX_SYMBOL[j]+str(i+2))
            worksheet.add_data_validation(dv)

    workbook.save(output_excel)

def highligh_former_manager(writer, tab):
    RECOMMAND_MANAGER_SYMBOL = ['D','F','H','J','L','N','P','R','T','V']
    workbook = writer.book
    worksheet = workbook[tab]
    redFill = PatternFill(start_color='FFA042', end_color='FFA042', fill_type='solid')
 
    for s in RECOMMAND_MANAGER_SYMBOL:
        col1 = worksheet[s]
        for i, cell in enumerate(col1):
            cell_value = cell.value
            if cell_value in former_manager:
                rule = Rule(type="cellIs", operator="equal", formula=[f'"{cell_value}"'], dxf=DifferentialStyle(fill=redFill))
                # rule = Rule(type="expression", operator="equal", formula=[f'"{cell_value}"'], dxf=DifferentialStyle(fill=redFill))
                worksheet.conditional_formatting.add(f'{s}{str(i+1)}', rule)

    workbook.save(output_excel)

def draw_color_for_similarity_score(writer, tab):
    SIMILARITY_SCORE_RANGE = '$E$2:$w$1000'
    workbook = writer.book
    worksheet = workbook[tab]
    rule = ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color="F9F900")
    worksheet.conditional_formatting.add(SIMILARITY_SCORE_RANGE, rule)
    workbook.save(output_excel)

search_and_analyze()