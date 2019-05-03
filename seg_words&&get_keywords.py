# -*- coding: utf-8 -*-
"""
Created on Fri May  3 17:51:02 2019

@author: Administrator
"""

import pandas as pd
import jieba
import jieba.analyse






#分词函数
def seg_words():
    global text
    #jieba.add_word('')  #可以往语料库中添加自定义词汇
    seg_list = jieba.cut(text, cut_all=False)      # 对读取的text分词,为精准模式
    return seg_list

#调用停用词表
def get_stopwords_list():
    stop_word_path = r'path of stopwordlist.txt'  #停用词表路径
    stopword_list = [sw.replace('\n','') for sw in open(stop_word_path).readlines()]  #实现按行读取停用词表
    return stopword_list


#过滤干扰词，需要用到停用词表
def word_filter(seg_list):
    stopword_list = get_stopwords_list()
    filter_list = []
    for word in seg_list:
        if not word in stopword_list:   #判断词是否在停用词表中
            filter_list.append(word)
    return filter_list

#处理数据集
def load_data():
    global text
    #调用已有函数处理数据集，处理后的每条数据仅保留非干扰词
    
    doc_list = []
    data = pd.read_excel(r'path + filename.xlsx').astype(str)   # 读取Excel数据
    line_num = len(data)     # 计算Excel中总记录数
    print(line_num)        # 显示总评论数
    
    counter=0  #counter计算行数
    while counter < line_num:
        text = data.iloc[counter,1]          # 读取counter行，第2列单元格内容
        seg_list = seg_words()
        filter_list = word_filter(seg_list)
        doc_list.append(filter_list)
        counter = counter + 1                          # counter+1，处理下一行数据
        #data.to_excel(r"C:\Users\Administrator\Desktop\dwjm\微博数据2.xlsx")    # 把结果保存到新Excel表中
    f = open(r"path + doc_list.txt","w",encoding='utf-8')   #创建doc_list的存放文件及路径
    f.write(str(doc_list))                         #写入文件
    return doc_list
    
    
    
    

    
#基于TF-IDF算法的关键词抽取
def get_tfidf():
    f = open(r"path of doc_list.txt","rb") #doc_list.txt的路径，打开文件
    f = f.read()
    keywords =jieba.analyse.extract_tags(f,topK=20,withWeight=True,allowPOS=()) 
#第一个参数是读取的文本，第二个参数是选取关键词数量，第三个关键词是是否显示权重，第四个关键词是选择显示词性标注，默认为空
    print("keywords by tfidf:")
# 输出抽取出的关键词
    for keyword in keywords:
        print(keyword[0],keyword[1]) #输出关键词和权重
        
#主函数
if __name__=="__main__":
    load_data()
    get_tfidf()