import urllib.request as req
import urllib
import os
import time
from urllib.parse import urljoin
import re
import time
import sys
import pandas as pd
from tqdm import tqdm
import docx

##################################################
#Step2: 段落を文ごとにカット
# →　エクセルに出力したあと、人力で整合的になるよう並び替え
####################main code#####################
#読み込みファイル
path = os.getcwd()
folder = path + "/2_make_corpus"

files = [
    f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))
]

#重複行は削除
df_corpus = pd.read_excel(path + "/corpus.xlsx")

#データフレームの準備
df_corp = pd.DataFrame()

#資料リストの準備
df_reflist = pd.read_excel(path + "/0_input/資料リスト.xlsx")

for file in tqdm(files):
    
    #読み込み
    filename =path + "/2_make_corpus/" + file
    df = pd.read_excel(filename)
    
    #スペース、NAN消し去り
    df = df.replace("　", "").replace(" ", "").replace('\n',"")
    df = df[["JP", "EN"]].dropna(how = "any")
    
    source = df_reflist[df_reflist["ファイル名"]==file.replace("_corp.xlsx","")].iloc[0,1]
    date  = df_reflist[df_reflist["ファイル名"]==file.replace("_corp.xlsx","")].iloc[0,2]
    df["出所"] = source
    df["日付"] = date
    
    #df_corpus = df
    df_corpus = pd.concat([df_corpus,df], axis = 0, ignore_index=True)
    
    #重複行は削除
    df_corpus = df_corpus.drop_duplicates()
    
    df_corpus.to_excel(path + "\corpus.xlsx", index = False, encoding = "shift-jis")
    
#     for p in range(pn):
        
#         #カット
#         para_JP = re.split(r'(?<=。)', str(df.iloc[p,0]))
#         df_para_JP = pd.DataFrame(para_JP)
        
#         para_EN = re.split(r'(?<=\. )', str(df.iloc[p,1]))
#         df_para_EN = pd.DataFrame(para_EN)
        
#         #スペースの削除
#         df_para_JP = df_para_JP.replace("　", "").replace(" ", "")
#         df_para_EN = df_para_EN.replace("　", "").replace(" ", "")
        
#         df_para_JP = df_para_JP[lambda df_para_JP: df_para_JP.iloc[:,0].str.len() > 0]
#         df_para_EN = df_para_EN[lambda df_para_EN: df_para_EN.iloc[:,0].str.len() > 0]
    
#         df_para = pd.concat([df_para_JP, df_para_EN], axis = 1, ignore_index=True)
#         df_para.columns = ["JP", "EN"]
#         df_para["paragraph"] = p
        
#         #結合
#         if p == 0:
#             df_out = df_para
#         else:
#             df_out = pd.concat([df_out,df_para], axis = 0, ignore_index=True)
            
#         #出力    
#         outname = file.replace("_para.xlsx", "")
#         df_out["file"] = outname
#         filename = path + "/2_make_corpus/" + outname + "_corp.xlsx"
#         df_out.to_excel(filename,index = False, encoding = "shift-jis")
            
    
        
# #     master = pd.concat([testdf, testdf_EN], axis=1, ignore_index=True)
# #     master.columns = ["JP", "EN"]
# #     outname = file.replace("_JP", "").replace(".docx", "")
# #     master["file"] = outname

# # master.to_excel(path + "/1_make_paragraph/" + outname + "_para.xlsx", \
#     #index = False, encoding = "shift-jis")