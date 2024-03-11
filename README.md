# Research Project Retrieval

## How to use
1. preprare 申請案件 excel, 統計 excel, 曾任委員 txt, put them into data folder (remember to change the project need to search in 統計)
2. run load_into_chroma_bge_manager.py, which will save embedding into chroma
3. run search_v3.py
4. the excusion result will be save to data folder
5. copy the excel formula in 委員推薦次數統計 tab to the generate one
6. replace filename in formula to '' empty string, which will change reference to current file
7. draw the graph by excel data
