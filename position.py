# -*- coding: utf-8 -*- 
import urllib.request, urllib.parse, urllib.error
import json
import hashlib
import csv
import xlwings as xw 


# 输出格式为json
output = 'json'
# 开发者平台获取的ak
ak = 'A1l2PXFLZ9eUElzprjeTfDN1NlcYw52B'
#开发者平台获取的sk
sk='7LFPvLbmz75To0bgFCx3ml9j1o7dLGvl'
# 目标地理位置,这里可以外部导入

a = ['庐江马槽']
''' # a=['北京','首都医大学','天坛医院','天通苑','德州','杭州','上海','北京大学','天津'] '''
# 打开保存位置
csv_obj = open('E:\code\data.csv\data.csv', 'w',newline='', encoding="GBK")
#写入title
csv.writer(csv_obj).writerow(["位置","x","y"])


# 进行爬取
for i in a:
    queryStr = '/geocoding/v3/?address={}&output=json&ak={}'.format(i,ak)
    #进行转码，safe为不转码的部分
    encodedStr = urllib.parse.quote(queryStr, safe="/:=&?#+!$,;'@()*[]")
    # 添加sk
    rawStr = encodedStr + sk
    # 算sn值，用于调用百度接口
    # 这里可以参看官方文档
    sn = (hashlib.md5(urllib.parse.quote_plus(rawStr).encode("utf8")).hexdigest())
    # 拼接url
    url = urllib.parse.quote("http://api.map.baidu.com" + queryStr + "&sn=" + sn, safe="/:=&?#+!$,;'@()*[]")
    # 目标请求
    req = urllib.request.urlopen(url)
    # 进行解码
    res = req.read().decode()
    # json转换为字典
    temp = json.loads(res)
    # 提取经度和纬度
    lng,lat=temp['result']['location']['lng'],temp['result']['location']['lat']
    # 写入csv文件
    csv.writer(csv_obj).writerow([i,lng,lat])


# 关闭csv文件
csv_obj.close()

