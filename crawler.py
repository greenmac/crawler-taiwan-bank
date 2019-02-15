# https://www.youtube.com/watch?v=-c5rrzjsN34
# http://idobest.pixnet.net/blog/post/35324197-%5B%E5%BF%83%E5%BE%97%5D-python%E7%88%AC%E8%9F%B2%E6%95%99%E5%AD%B82018-%E6%8A%93%E5%8F%96%E5%8F%B0%E9%8A%80%E7%89%8C%E5%91%8A%E5%8C%AF%E7%8E%87
# https://rate.bot.com.tw/xrt?Lang=zh-TW
import pandas as pd
from pymongo import MongoClient
import datetime
import time

start_time = time.time()
dfs = pd.read_html("https://rate.bot.com.tw/xrt?Lang=zh-TW")
currency = dfs[0]
currency_fix = currency.ix[:, 0:5]
# 轉成Unicode比較不會有中文亂碼的問題
currency_fix.columns = [u'幣別', u'現金匯率-本行買入', u'現金匯率-本行賣出', u'即期匯率-本行買入', u'即期匯率-本行賣出']
currency_fix[u'幣別'] = currency_fix[u'幣別'].str.extract('\((\w+)\)')
add_time = datetime.datetime.now()
currency_fix[u'更新時間'] = add_time
# print(currency_fix)

# 存檔
# currency_fix.to_excel('currency.xlsx') # 存成excel
# currency_fix.to_csv('currency.csv') # 存成csv

# mongodb
client = MongoClient()
database = client['taiwan_bank']
collection = database['exchange_rate']

records = currency_fix.to_dict('records')
collection.insert_many(records)
end_time = time.time()
cost_time = end_time - start_time
print("cost_time:", cost_time)