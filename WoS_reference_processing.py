"""
Created on Tue Aug 10 15:07:01 2021

@author: yonniye
"""

import pandas as pd
import glob
import requests
import time
from bs4 import BeautifulSoup
from pygtrans import Translate

def get_IF_from_name(name):
    headers={
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.22 Safari/537.36 SE 2.X MetaSr 1.0'
    }
    name.replace(' ', '+')
    address = "http://sci.justscience.cn/?q=" + str(name) + "&sci=1"
    f = requests.get(address, headers = headers, timeout=(3,7))
    soup = BeautifulSoup(f.text, features="lxml")
    tds = soup.find_all(height="30")
    for i in range(0, len(tds), 8):
        if tds[i+1].text.upper() == name.upper():
            time.sleep(1)
            return tds[i+7].text

def get_IF_from_issn(issn):
    headers={
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/49.0.2623.22 Safari/537.36 SE 2.X MetaSr 1.0'
    }
    address = "http://sci.justscience.cn/?q=" + str(issn) + "&sci=1"
    f = requests.get(address, headers = headers)
    soup = BeautifulSoup(f.text, features="lxml")
    return soup.find_all(height="30")[-1].text  # 结果唯一

def main():
    excel_list = glob.glob(r"D:\Desktop\list\*.xls")
    
    print(f"共要处理 {len(excel_list)} 个文件")
    for exl in excel_list:
        print(f'正在处理 {exl}')
        df = pd.read_excel(exl)
        max_column=67 # wos全导出后共67个字段
        column_saved = ['Author Full Names', 'Article Title', 'Source Title', 'Document Type', 'Author Keywords', 'Keywords Plus', 'Abstract', 'Times Cited, WoS Core', 'ISSN', 'Publication Year', 'DOI']
        
        # 保留所需列
        for col in df.columns.values:
            if col not in column_saved:
                df.drop(labels=col, axis=1, inplace=True)
                
        # 添加所需列
        df.insert(3, 'IF(2021)', value='0.0', allow_duplicates=False)
        df.insert(8, 'Translation', value='', allow_duplicates=False)
        
        nrows = df.shape[0]
        
        IF_index = list(df.columns).index('IF(2021)')
        Trans_index = list(df.columns).index('Translation')
        
        client = Translate()
        
        for i in range(nrows):
            df.iloc[i, Trans_index] = client.translate(df['Abstract'][i]).translatedText
            
            if len(str(df['ISSN'][i]))==9:
                try:
                    df.iloc[i, IF_index] = get_IF_from_issn(df['ISSN'][i])
                except:
                    df.iloc[i, IF_index] = "nan"
            else:
                try:
                    df.iloc[i, IF_index] = get_IF_from_name(df['Source Title'][i])
                except:
                    df.iloc[i, IF_index] = "nan"
                
        df.to_excel(".\postprocessing\output_"+exl.split("\\")[-1].split(".")[0]+".xlsx")
        
        print(exl.split("\\")[-1] + "处理完成！")
        time.sleep(4)
        
if __name__ == "__main__":
    main()