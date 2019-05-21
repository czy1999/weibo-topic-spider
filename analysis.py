from aip import AipNlp
import pandas as pd
import numpy as np
import time



# 此处输入baiduAIid
APP_ID = ''
API_KEY = ''
SECRET_KEY = ''


client = AipNlp(APP_ID, API_KEY, SECRET_KEY)

data = pd.read_excel('teamwe.xls',encoding='utf-8')

def isPostive(text):
    try:
        if client.sentimentClassify(text)['items'][0]['positive_prob']>0.5:
            return "积极"
        else:
            return "消极"
    except:
        return "积极"


data = pd.read_excel('mlxg.xls',encoding='utf-8')

aa = []
count = 1
for i in data['微博内容']:
    aa.append(isPostive(i))
    count+=1
    print(count)


def fenlei(text):
    for j in cz:
        if j in text:
            return "创作"
    for i in xf:
        if i in text:
            return "消费"
    for k in gj:
        if k in text:
            return "攻击"
    return "其他"        
    


xf = ['抽奖',"抽一个","抽一位","买","通贩"]
cz = ["画","实物","返图","合集","摸鱼","漫","自制","攻略","授权","草稿","绘"]
gj = ["hz","狗粉丝","狗女儿"]

b= []
for ix in data['微博内容']:
    b.append(fenlei(ix))
