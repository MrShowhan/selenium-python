import xlrd
import xlwt
from xlutils.copy import copy
from selenium.webdriver.common.by import By  # 按照什么方式查找，By.ID,By.CSS_SELECTOR
from selenium.webdriver.common.keys import Keys  # 键盘按键操作
from selenium.webdriver.support import expected_conditions as EC  # 和下面WebDriverWait一起用的
from selenium.webdriver.support.wait import WebDriverWait  # 等待页面加载某些元素
from selenium.webdriver import Chrome,ChromeOptions
from lxml import etree
import os
import re
import time

def start_driver():
    driver = Chrome("G:\chromedriver.exe")
    """
    需要设置navigator的webdriver值为Chrome才行，undefined也会被屏蔽
    """
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
      "source": """
        Object.defineProperty(navigator, 'webdriver', {
          get: () => Chrome         
        })
      """
    })
    return driver

def get_html(driver,keyword):
    wait = WebDriverWait(driver,10)
    driver.get('http://ftba.nmpa.gov.cn:8181/ftban/fw.jsp')
    input_tag=wait.until(EC.presence_of_element_located((By.ID,"searchtext")))  #定位查询输入窗口
    input_tag.send_keys(keyword)   #设置需要搜索的产品名称
    input_tag.send_keys(Keys.ENTER) #使用回车进行确认查询
    page = driver.page_source   #获取网页源代码
    if re.search(r'抱歉，未检索到相关数据!',page): #判断查询结果是否正确
        print('抱歉，未检索到相关数据!')
        driver.close()  #关闭浏览器
        return None
    else:
        return page

def total_pages(page):      #获取总页码
    html = etree.HTML(page)
    ul = html.xpath('//li[@class="xl-nextPage"]/preceding-sibling::li/text()')  #定位”下一页“元素的前兄弟元素
    pages = ul[-1]  #最后一页页码元素的位置
    print('一共有{}页'.format(pages))
    return int(pages)

def get_data(page):     #使用XPATH从网页源代码提取目标数据
    html = etree.HTML(page)
    date = html.xpath('//ul[@id="gzlist"]/li/i/text()') #提取日期
    link = html.xpath('//ul[@id="gzlist"]/li/dl/a/@href')   #提取连接
    dl_title = html.xpath('//ul[@id="gzlist"]/li/dl/a/text()')  #提取产品标题
    ol_title = html.xpath('//ul[@id="gzlist"]/li/ol/a/text()')  #备案号
    company = html.xpath('//ul[@id="gzlist"]/li/p/text()')  #公司名称
    total_list = []
    for i in range(len(date)):
        list =[]
        list.append(dl_title[i])
        list.append(ol_title[i])
        list.append(company[i])
        list.append(date[i])
        list.append(link[i])
        total_list.append(list) #使用列表形式整合数据返回
    return total_list

def next_page(driver,total_page):   #翻页
    for i in range(total_page-1):
        driver.find_element(By.XPATH,'//li[@class="xl-nextPage"]').click()  #点击下一页
        page = driver.page_source
        yield page  #使用生成器返回每一页源代码

def write_excel_xls(path, sheet_name, value):
    workbook = xlwt.Workbook()  # 新建一个工作簿
    sheet = workbook.add_sheet(sheet_name)  # 在工作簿中新建一个表格
    for i in range(0, len(value)):
        sheet.write(0, i, value[i])
    workbook.save(path)  # 保存工作簿
    print("xls表格创建成功！")


def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i + rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿
    print("【追加】写入数据成功！")



if __name__ == '__main__':
    while True:
        keyword = input('请输入查询内容：')
        strat_time = time.time()
        driver = start_driver()
        page = get_html(driver,keyword)
        if page !=None:
            break
    pages = total_pages(page)
    data = get_data(page)
    if not os.path.exists(r'.\{}.xls'.format(keyword)): #判断是否需要新建xls
        first_row =['产品名称','备案号','公司名称','日期','详情连接']    #设置xls表格中第一行的内容
        write_excel_xls(r'.\{}.xls'.format(keyword),'sheet1',first_row)
    write_excel_xls_append(r'.\{}.xls'.format(keyword),data)    #在已有数据的xls中追加写入信息
    nextpage = next_page(driver,pages)
    for page in nextpage:
        data = get_data(page)
        write_excel_xls_append(r'.\{}.xls'.format(keyword), data)   #翻页后在已有数据的xls中追加写入信息
    driver.close()
    print('一共用时{}秒'.format(time.time()-strat_time)) #耗时

