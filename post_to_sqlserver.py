import pandas as pd
import os
import sys
import datetime
import requests
import json
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

searchterm_path = ""
start_time = datetime.datetime.now()

keyword_file_count = 0
for file in os.listdir("./searchterm"):
    if (".xlsx" in file):
        keyword_file_count += 1
        searchterm_path = f"./searchterm/{file}"

if (keyword_file_count != 1):
    print("there should be 1 and only 1 .xlsx file in searchterm folder")
    sys.exit(0)

print("start processing...")

print("开始，过去了")
print(datetime.datetime.now() - start_time)

keyword_df = pd.read_excel(searchterm_path)

print("读取 搜索词，过去了")
print(datetime.datetime.now() - start_time)

for num in range(0, keyword_df.index.size):
    # print(f"analysing record {num}...")

    row = keyword_df.loc[num]

    url = 'http://localhost:5000/searchterm/upload'
    st = {
        "market": row["部门"],
        "name": row["搜索词"],
        "rank": int(row["搜索频率排名"]),  #json 不识别 NumPy int64 类型，转换为 int 
        "firstASIN": row["#1 已点击的 ASIN"],
        "firstASINName": row["#1 商品名称"],
        "firstASINClickShare": row["#1 点击共享"],
        "firstASINConvertShare": row["#1 转化共享"],
        "secondASIN": row["#2 已点击的 ASIN"],
        "secondASINName": row["#2 商品名称"],
        "secondASINClickShare": row["#2 点击共享"],
        "secondASINConvertShare": row["#2 转化共享"],
        "thirdASIN": row["#3 已点击的 ASIN"],
        "thirdASINName": row["#3 商品名称"],
        "thirdASINClickShare": row["#2 点击共享"],
        "thirdASINConvertShare": row["#3 转化共享"]
    }

    res = requests.post(
        url,
        data=json.dumps(st),
        verify=False,
        headers={"Content-Type": "application/json; charset=UTF-8"})
    # print(res.content)

print("完成计算，过去了")
print(datetime.datetime.now() - start_time)
