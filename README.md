# selenium-for-sinaweibo-spyder
python3.6+selenium+Chrome to crawl data from sinaweibo
这里使用的是selenium模拟微博手机版登录，搜索特定的某个话题并爬取搜索结果的微博内容、用户名、点赞数、转发数和发布时间，最后将数据逐行存入excel文件。目的最初是用于了解在校学生对所在高校的评价，需要做情感分析。
后期可能更新中文分词、关键词抽取和情感分析的内容（NLP）。
By the way,记录一下遇到的小问题：对excel的操作我使用的是openpyxl的库，这个库支持对xlsx文件的操作，xls操作会报错。
爬虫萌新程序猿，欢迎大家提出批评建议。
