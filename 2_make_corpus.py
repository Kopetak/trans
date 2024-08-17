import os
import re
import pandas as pd
from tqdm import tqdm

##################################################
#Step2: 段落を文ごとにカット
# →　エクセルに出力したあと、人力で整合的になるよう並び替え
####################main code#####################
#読み込みファイル
path = os.getcwd()
folder = path + "/1_make_paragraph"
files = [
    f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))
]

#データフレームの準備
df_corp = pd.DataFrame()
# df_corp.columns = ["JP", "EN", "para", "file"]

for file in tqdm(files):
    
    #読み込み
    filename =path + "/1_make_paragraph/" + file
    df = pd.read_excel(filename)
    df = df[["JP", "EN"]].dropna(how = "all")
    
    pn = len(df)  # パラグラフ数
    
    for p in range(pn):
        
        #カット
        para_JP = re.split(r'(?<=。)', str(df.iloc[p,0]))
        df_para_JP = pd.DataFrame(para_JP)
        
        EN_split = re.compile(r'(?<!U\.S\. )(?<!I\. )(?<!II\. )(?<!III\. )(?<=\. )')
        para_EN = re.split(EN_split, str(df.iloc[p,1]))
        df_para_EN = pd.DataFrame(para_EN)
        
        #スペースの削除
        df_para_JP = df_para_JP.replace("　", "").replace(" ", "")
        df_para_EN = df_para_EN.replace("　", "").replace(" ", "")
        
        df_para_JP = df_para_JP[lambda df_para_JP: df_para_JP.iloc[:,0].str.len() > 0]
        df_para_EN = df_para_EN[lambda df_para_EN: df_para_EN.iloc[:,0].str.len() > 0]
    
        df_para = pd.concat([df_para_JP, df_para_EN], axis = 1, ignore_index=True)
        df_para.columns = ["JP", "EN"]     
        bl_row = {'JP': ' ', 'EN': ' '}
        df_para = df_para.append(bl_row, ignore_index=True)
        df_para["paragraph"] = p
        
        #結合
        if p == 0:
            df_out = df_para
        else:
            df_out = pd.concat([df_out,df_para], axis = 0, ignore_index=True)
            
        #出力    
        outname = file.replace("_para.xlsx", "")
        df_out["file"] = outname
        filename = path + "/2_make_corpus/ori/" + outname + "_corp.xlsx"
        df_out.to_excel(filename,index = False, encoding = "shift-jis")
            
    
        
#     master = pd.concat([testdf, testdf_EN], axis=1, ignore_index=True)
#     master.columns = ["JP", "EN"]
#     outname = file.replace("_JP", "").replace(".docx", "")
#     master["file"] = outname

# master.to_excel(path + "/1_make_paragraph/" + outname + "_para.xlsx", \
    #index = False, encoding = "shift-jis")