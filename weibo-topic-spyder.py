import time
import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os
import excelSave as save

# 用来控制页面滚动
def Transfer_Clicks(browser):
    time.sleep(5)
    try:
        browser.execute_script("window.scrollBy(0,document.body.scrollHeight)", "")
    except:
        pass
    return "Transfer successfully \n"

#判断页面是否加载出来
def isPresent():
    temp =1
    try: 
        driver.find_elements_by_css_selector('div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
    except:
        temp =0
    return temp

#插入数据
def insert_data(elems,path,name,yuedu,taolun):
    for elem in elems:
        workbook = xlrd.open_workbook(path)  # 打开工作簿
        sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
        worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
        rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数       
        rid = rows_old
        #用户名
        weibo_username = elem.find_elements_by_css_selector('h3.m-text-cut')[0].text
        weibo_userlevel = "普通用户"
        #微博等级
        try: 
            weibo_userlevel_color_class = elem.find_elements_by_css_selector("i.m-icon")[0].get_attribute("class").replace("m-icon ","")
            if weibo_userlevel_color_class == "m-icon-yellowv":
                weibo_userlevel = "黄v"
            if weibo_userlevel_color_class == "m-icon-bluev":
                weibo_userlevel = "蓝v"
            if weibo_userlevel_color_class == "m-icon-goldv-static":
                weibo_userlevel = "金v"
            if weibo_userlevel_color_class == "m-icon-club":
                weibo_userlevel = "微博达人"     
        except:
            weibo_userlevel = "普通用户"
        #微博内容
        weibo_content = elem.find_elements_by_css_selector('div.weibo-text')[0].text
        shares = elem.find_elements_by_css_selector('i.m-font.m-font-forward + h4')[0].text
        comments = elem.find_elements_by_css_selector('i.m-font.m-font-comment + h4')[0].text
        likes = elem.find_elements_by_css_selector('i.m-icon.m-icon-like + h4')[0].text

        #发布时间
        weibo_time = elem.find_elements_by_css_selector('span.time')[0].text
        '''
        print("用户名："+ weibo_username + "|"
              "微博等级："+ weibo_userlevel + "|"
              "微博内容："+ weibo_content + "|"
              "转发："+ shares + "|"
              "评论数："+ comments + "|"
              "点赞数："+ likes + "|"
              "发布时间："+ weibo_time + "|"
              "话题名称" + name + "|" 
              "话题讨论数" + yuedu + "|"
              "话题阅读数" + taolun)
        '''
        value1 = [[rid, weibo_username, weibo_userlevel,weibo_content, shares,comments,likes,weibo_time,keyword,name,yuedu,taolun],]
        print("当前插入第%d条数据" % rid)
        save.write_excel_xls_append_norepeat(book_name_xls, value1)
#获取当前页面的数据
def get_current_weibo_data(elems,book_name_xls,name,yuedu,taolun,maxWeibo):
    #开始爬取数据
        before = 0 
        after = 0
        n = 0 
        timeToSleep = 100
        while True:
            before = after
            Transfer_Clicks(driver)
            time.sleep(3)
            elems = driver.find_elements_by_css_selector('div.card.m-panel.card9')
            print("当前包含微博最大数量：%d,n当前的值为：%d, n值到5说明已无法解析出新的微博" % (len(elems),n))
            after = len(elems)
            if after > before:
                n = 0
            if after == before:        
                n = n + 1
            if n == 5:
                print("当前关键词最大微博数为：%d" % after)
                insert_data(elems,book_name_xls,name,yuedu,taolun)
                break
            if len(elems)>maxWeibo:
                print("当前微博数以达到%d条"%maxWeibo)
                insert_data(elems,book_name_xls,name,yuedu,taolun)
                break
            '''
            if after > timeToSleep:
                print("抓取到%d多条，插入当前新抓取数据并休眠5秒" % timeToSleep)
                timeToSleep = timeToSleep + 100
                insert_data(elems,book_name_xls,name,yuedu,taolun) 
                time.sleep(5) 
            '''


#爬虫运行 
def spider(username,password,driver,book_name_xls,sheet_name_xls,keyword,maxWeibo):
    
    #创建文件
    if os.path.exists(book_name_xls):
        print("文件已存在")
    else:
        print("文件不存在，重新创建")
        value_title = [["rid", "用户名称", "微博等级", "微博内容", "微博转发量","微博评论量","微博点赞","发布时间","搜索关键词","话题名称","话题讨论数","话题阅读数"],]
        save.write_excel_xls(book_name_xls, sheet_name_xls, value_title)
    
    #加载驱动，使用浏览器打开指定网址  
    driver.set_window_size(452, 790)
    driver.get("https://passport.weibo.cn/signin/login")
    print("开始自动登陆，若出现验证码手动验证")
    time.sleep(3)

    elem = driver.find_element_by_xpath("//*[@id='loginName']");
    elem.send_keys(username)
    elem = driver.find_element_by_xpath("//*[@id='loginPassword']");
    elem.send_keys(password)
    elem = driver.find_element_by_xpath("//*[@id='loginAction']");
    elem.send_keys(Keys.ENTER)  
    print("暂停20秒，用于验证码验证")
    time.sleep(20)

    '''
    # 添加cookie
    cookie = []
    for ix in cookie:
        driver.add_cookie(ix)
    driver.get("https://m.weibo.cn")
    '''
        
        
    while 1:  # 循环条件为1必定成立
        result = isPresent()
        print ('判断页面1成功 0失败  结果是=%d' % result )
        if result == 1:
            elems = driver.find_elements_by_css_selector('div.line-around.layout-box.mod-pagination > a:nth-child(2) > div > select > option')
            #return elems #如果封装函数，返回页面
            break
        else:
            print ('页面还没加载出来呢')
            time.sleep(20)

    time.sleep(2)

    #搜索关键词
    elem = driver.find_element_by_xpath("//*[@class='m-text-cut']").click();
    time.sleep(2)
    elem = driver.find_element_by_xpath("//*[@type='search']");
    elem.send_keys(keyword)
    elem.send_keys(Keys.ENTER)
    time.sleep(5)
    
    # elem = driver.find_element_by_xpath("//*[@class='box-left m-box-col m-box-center-a']")
    # 修改为点击超话图标进入超话，减少错误
    elem = driver.find_element_by_xpath("//img[@src ='http://simg.s.weibo.com/20181009184948_super_topic_bg_small.png']")
    elem.click()
    elem.click()
    print("超话链接获取完毕，休眠2秒")
    time.sleep(2)
    yuedu_taolun = driver.find_element_by_xpath("//*[@id='app']/div[1]/div[1]/div[1]/div[4]/div/div/div/a/div[2]/h4[1]").text
    yuedu = yuedu_taolun.split("　")[0]
    taolun = yuedu_taolun.split("　")[1]
    time.sleep(2)
    name = keyword
    shishi_element = driver.find_element_by_xpath("//*[@class='scroll-box nav_item']/ul/li/span[text()='帖子']")
    driver.execute_script('arguments[0].click()',shishi_element) 
    get_current_weibo_data(elems,book_name_xls,name,yuedu,taolun,maxWeibo) #爬取实时
    time.sleep(2)

    
if __name__ == '__main__':
    username = "" #你的微博登录名
    password = "" #你的密码
    driver = webdriver.Chrome()#你的chromedriver的地址
    book_name_xls = "test.xls" #填写你想存放excel的路径，没有文件会自动创建
    sheet_name_xls = '微博数据' #sheet表名
    maxWeibo = 5000 #设置最多多少条微博
    keywords = ["肺炎",] # 此处可以设置多个超话关键词 
    for keyword in keywords:
        spider(username,password,driver,book_name_xls,sheet_name_xls,keyword,maxWeibo)
