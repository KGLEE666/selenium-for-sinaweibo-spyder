# -*- coding: utf-8 -*-
"""
Created on Sat Mar 23 21:01:55 2019

@author: Administrator
"""

from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys 
import selenium.webdriver.support.ui as ui 
from openpyxl import workbook
from openpyxl import load_workbook
import re              
import os      
import sys    
import codecs      
import urllib   


options = webdriver.ChromeOptions()
#设置请求头,参考网址：https://httpbin.org/get?show_env=1
user_agent = (
    "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36"
    )
options.add_argument('user-agent=%s'%user_agent)

#禁止加载图片，提高速度
options.add_argument('blink-settings=imagesEnabled=false')

#调用chrome浏览器
driver = webdriver.Chrome(executable_path='C:\Program Files (x86)\Google\Chrome\Application\chromedriver')
driver.set_window_size(1366, 768)
wait = ui.WebDriverWait(driver,10)

#全局变量 txt文件操作读写信息，因为后来用excel所以注释掉 
#infofile = codecs.open("SinaWeibo_Info.txt", 'a', 'utf-8')  

#登录微博手机版
def LoginWeibo(username, password):  
    try:  
        #**********************************************************************  
        # 登录函数，直接访问driver.get("http://weibo.cn/5824697471")会跳转到登陆页面 用户id       
        #**********************************************************************  
          
        #输入用户名/密码登录  
        print(u'准备登陆Weibo.cn网站...')  
        
        driver.get("https://passport.weibo.cn/signin/login") 
        time.sleep(3)
        elem_user = driver.find_element_by_id("loginName")  
        elem_user.send_keys('18688774843') #用户名  
        elem_pwd = driver.find_element_by_id("loginPassword")  
        elem_pwd.send_keys('951219KG')  #密码  
        #elem_rem = driver.find_element_by_name("remember")  
        #elem_rem.click()             #记住登录状态  
  
        #重点: 暂停时间输入验证码  
        #pause(millisenconds)  
        time.sleep(5)
        elem_signin = driver.find_element_by_id("loginAction")
        elem_signin.click()
        #点击登陆  
        time.sleep(10)
        #获取Coockie 推荐 http://www.cnblogs.com/fnng/p/3269450.html  
        print(driver.current_url)  
        print(driver.get_cookies())   #获得cookie信息 dict存储  
        print(u'输出Cookie键值对信息:')   
        for cookie in driver.get_cookies():   
            #print cookie  
            for key in cookie:  
                print(key, cookie[key])   
                      
        #driver.get_cookies()类型list 仅包含一个元素cookie类型dict  
        print(u'登陆成功...')   
        time.sleep(3)
        
        
        
    except Exception as e: 
        print('e.message:\t', e) 
    finally:    
        print(u'End LoginWeibo!\n\n')
        
#*******************************************************************************
#第二步：访问http://s.weibo.com/页面搜索热点信息
#*******************************************************************************
def do_search(key):
    try:
        driver.get("http://s.weibo.cn/")
        print(u'搜索热点主题：', key)
        
        #输入主题并点击搜索
        item_inp = driver.find_element_by_link_text("搜索")
        time.sleep(2)
        item_inp.click()
        time.sleep(5)
        item_search = driver.find_element_by_name("keyword")
        #item_inp.click() 点击搜索框，好像没用
        item_search.clear()           #清空输入框
        time.sleep(2)
        item_search.send_keys(key)    #输入搜索内容
        time.sleep(2)
        item_search.send_keys(Keys.ENTER)    #回车搜索
        time.sleep(5)
            
    except Exception as e:      
        print("Error: ",e)
        #GetComment(key)    #在考虑是否报错时再调用这个方法
    finally:
        print(u'do searching\n\n')
#********************************************************************************
#                  第三步: 
#                  爬取微博内容，注意评论翻页的效果和微博的数量
#********************************************************************************    
 
def getContent():
    try:
        #global infofile       #全局文件变量
        global page #声明全局变量
        while page < 73:#搜索到的内容共有72页
            user_id = driver.find_elements_by_class_name("nk")  #获取用户名元素
            content = driver.find_elements_by_class_name("ctt") #获取内容元素
            zan = driver.find_elements_by_partial_link_text("赞")#获取点赞数元素
            zhuanfa = driver.find_elements_by_partial_link_text("转发")#获取转发数元素
            moment = driver.find_elements_by_class_name("ct") #获取发布时间元素
            
                
            print(u'长度:', len(content))
            i = 0
            while i < len(content):
                print('用户名：')
                print(user_id[i].get_attribute('textContent'),'\n')
                print('微博信息:')
                print(content[i].get_attribute('textContent'),'\n')  #仅获取文本
                zan1 = str(zan[i].get_attribute('textContent'))  #zan返回的是列表，这里取列表的第i个元素并转为字符串
                zan_num = re.sub("\D","",zan1) #正则表达式，仅保留"赞【10】"中的数字10
                print('点赞：' + zan_num,'\n')
                zhuanfa1 = str(zhuanfa[i].get_attribute('textContent'))
                zhuanfa_num = re.sub("\D","",zhuanfa1) #仅保留“转发【10】“中的数字10
                print('转发：' + zhuanfa_num,'\n')
                moment1 = str(moment[i].get_attribute('textContent')).split('来') #'2016年4月15日 来自微博手机版'，只取前面的日期，不要‘来自’的内容
                f_moment = moment1[0] 
                print('发布时间：' + f_moment,'\n')
                ws.append([user_id[i].get_attribute('textContent'),content[i].get_attribute('textContent'),zan_num,zhuanfa_num,f_moment])   #写入ws工作表`
                #infofile.write(u'微博信息:\r\n')     #写入txt
                #infofile.write(content[i].get_attribute('textContent') + '\r\n')
                i = i + 1
                time.sleep(2)
            driver.implicitly_wait(10)   #设置隐式等待，如果没有翻页元素，等待10s
            page_move = driver.find_element_by_link_text("下页")   #翻页
            page_move.click()
            print('页码：'+ str(page))
            page = page + 1
            time.sleep(20)
        
        
        
        
    except Exception as e:      
        print("Error: ",e)
        
    finally:
        wb.save('filename.xlsx')   #保存已爬取的数据到excel
        con_load()                #开启续传
        print(u'oops!\n\n')
        

#*******************************************************************************
#断点续传逻辑
#*******************************************************************************
def con_load():
    try:
        do_search(key)    
        item_page = driver.find_element_by_name("page")
        item_page.send_keys(page)               #从断点页码处重新开始
        time.sleep(2)
        item_submit = driver.find_element_by_xpath("//*[@id='pagelist']/form/div/input[3]")
        item_submit.click()
        time.sleep(2)
        print('从第' +  str(page) + '页开始')
        getContent()
        
    except Exception as e:      
        print("Error: ",e)
        
    finally:
        print(u'续传失败!\n\n')
#*******************************************************************************
#                                主函数
#*******************************************************************************
    
if __name__ == '__main__':
 
    #定义变量
    username = ''             #输入你的用户名
    password = ''               #输入你的密码
    global page #创建页码
    page = 1
    #操作函数
    LoginWeibo(username, password)       #登陆微博
    wb = workbook.Workbook()  # 创建Excel对象
    ws = wb.active  # 获取当前正在操作的表对象
    # 往表中写入标题行,以列表形式写入！
    ws.append(['用户名','内容', '点赞数', '转发数','发布时间'])
    #搜索热点微博 爬取评论
    key = u'' #输入搜索关键词
    do_search(key)
    getContent()
    wb.save('filename.xlsx')  # 存入所有信息后，保存为filename.xlsx
    wb.close()
    #infofile.close()
    