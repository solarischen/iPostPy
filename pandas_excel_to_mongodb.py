# 引入模块
import pandas as pd
from pymongo import MongoClient
import json
import os
import datetime

start_time = datetime.datetime.now()

# 连接数据库 - 获取数据库 - 获取集合
mongo_client = MongoClient('mongodb://127.0.0.1:27017')
iFind_db = mongo_client["iFind"]
print(f"connecto iFInd database")

# 遍历文件, 先检查名字是否重复，然后执行计算
for filename in os.listdir("searchterm/"):
    print(f"start analyzing {filename}")
    # 连接到对应的集合
    searchterm_collection = iFind_db["SearchTerm"]
    print("connect to collection " + searchterm_collection.name)

    # 执行读取-写入操作
    with pd.ExcelFile(f"searchterm/{filename}") as xlsx:

        # 读取为 dataframe
        searchterms_frame = pd.read_excel(xlsx,
                                          converters={
                                              '搜索频率排名': int,
                                              '#1 点击共享': str,
                                              '#1 转化共享': str,
                                              '#2 点击共享': str,
                                              '#2 转化共享': str,
                                              '#3 点击共享': str,
                                              '#3 转化共享': str
                                          })
        print(f"read searchterm/{filename} to dataframe")
        # 重命名列
        searchterms_frame = searchterms_frame.rename(
            columns={
                "部门": "Market",
                "搜索词": "Name",
                "搜索频率排名": "Rank",
                "#1 已点击的 ASIN": "FirstASIN",
                "#1 商品名称": "FirstASINName",
                "#1 点击共享": "FirstASINClickShare",
                "#1 转化共享": "FirstASINConvertShare",
                "#2 已点击的 ASIN": "SecondASIN",
                "#2 商品名称": "SecondASINName",
                "#2 点击共享": "SecondASINClickShare",
                "#2 转化共享": "SecondASINConvertShare",
                "#3 已点击的 ASIN": "ThirdASIN",
                "#3 商品名称": "ThirdASINName",
                "#3 点击共享": "ThirdASINClickShare",
                "#3 转化共享": "ThirdASINConvertShare"
            })

        print(f"total {searchterms_frame.count} records")
        # 转换为 json
        data = json.loads(searchterms_frame.T.to_json()).values()
        print("convert to json, and ready to save")

        # 插入数据到数据库
        searchterm_collection.insert_many(data)
        print(f"finished saving {searchterm_collection.name} to database...")
        print("next...")

# 结束
print("完成计算，过去了")
print(datetime.datetime.now() - start_time)
