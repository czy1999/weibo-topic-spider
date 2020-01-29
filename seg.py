# -*- coding: utf-8 -*-

import jieba
from collections import Counter
import pandas as pd



def save_seg(filename,cnt):
    f_out = open(filename, 'w+')
    result = cnt.most_common(100)
    for ix in result:
        f_out.write(ix[0]+"\t出现次数："+str(ix[1])+"\n")



STOPWORDS = [u'的',u' ',u'\n',u'他', u'地', u'得', u'而', u'了', u'在', u'是', u'我', u'有', u'和', u'就',  u'不', u'人', u'都', u'一', u'一个', u'上', u'也', u'很', u'到', u'说', u'要', u'去', u'你',  u'会', u'着', u'没有', u'看', u'好', u'自己', u'这']
PUNCTUATIONS = [u'。',u'#', u'，', u'“', u'”', u'…', u'？', u'！', u'、', u'；', u'（', u'）']


# 需要进行分词的文件
wj = ['mlxg','IG+rng','igbanlan','edg','uzi','teamwe','theshy','英雄联盟','jackeylove']

cnt = Counter()

for file in wj:
    data = pd.read_excel(file+'.xls',encoding='utf-8') 
    for l in data['微博内容'].astype(str):
        seg_list = jieba.cut(l)
        for seg in seg_list:
            if seg not in STOPWORDS and seg not in PUNCTUATIONS and seg not in wj:
                cnt[seg] = cnt[seg] + 1


    save_seg("seg_result/"+file+".txt",cnt) # 保存文件

