from selenium import webdriver
from selenium.webdriver.common.by import By
import xlwt
from time import sleep




def main():
    baseurl="https://www.bilibili.com/v/popular/rank/all"
    datalist = getData(baseurl)
    savepath = "bilibili_rank.xls"
    saveData(datalist, savepath)
    
# get data from url and return datalist
def getData(baseurl):
    print("grabbing...")
    datalist = []
    browser = webdriver.Chrome()
    browser.get(baseurl)
    sleep(10)
    elements = browser.find_elements(By.CLASS_NAME, 'content')
    for elem in elements:
        data =[]
        findRank = elem.find_element(By.TAG_NAME, 'i').text
        findTitle = elem.find_element(By.CLASS_NAME, 'title').text
        findAuthor = elem.find_element(By.XPATH, './div[2]/div/a/span').text
        findView = elem.find_element(By.XPATH, './div[2]/div/div/span[1]').text
        data.append(findRank)
        data.append(findTitle)
        data.append(findAuthor)
        data.append(findView)
        print(data)
        datalist.append(data)
    browser.close()
    browser.quit()
    print("===grabed===")
    return datalist

    
def saveData(datalist, savepath):
    print("saving...")
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet('bilibili_top100', cell_overwrite_ok=True)
    col = ['rank', 'title', 'uploader', 'views']
    for i in range(0,4):
        sheet.write(0, i, col[i])
    if len(datalist) == 100:
        for i in range(0, 100):
            for j in range(0, 4):
                sheet.write(i+1, j, datalist[i][j])
    book.save(savepath)
    print("===saved===")

main()