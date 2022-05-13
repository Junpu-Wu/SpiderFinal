# 姓名：吴均浦
# 学号：2020310621287
# 爬取内容：亚马逊Kindle Oasis产品评论
import urllib.error
import requests
import ssl
import re
import bs4
from bs4 import BeautifulSoup
from urllib import request
import xlwt
import time


class Spider:
    def __init__(self, baseUrl, filename1):
        self.baseUrl = baseUrl
        self.dataList = []
        self.fname = filename1

    def gethtml(self, url):  # 网页请求
        head = {  # 请求头
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36 Edg/99.0.1150.55"
        }
        html = ""
        context = ssl._create_unverified_context()  # 设置证书
        try:
            response = requests.get(url, headers=head)
            text = response.status_code  # 获取并打印状态码
            print("status:{}".format(text))
            req = request.Request(url, headers=head)
            response2 = request.urlopen(req, context=context)
            html = response2.read().decode("utf-8")
        except urllib.error.URLError as e:
            if hasattr(e, "code"):
                print(e.code)
            if hasattr(e, "reason"):
                print(e.reason)
        return html

    def getData(self):
        for i in range(1, 351):
            print("Page:", i)
            url = self.baseUrl + str(i)
            html = self.gethtml(url)
            bs = BeautifulSoup(html, "html.parser")
            for item in bs.find_all('div', class_="a-section celwidget"):  # 以下为提取信息
                ranking = ""
                date = ""
                usefulNum = ""
                comment = ""  # 以上为初始化元素类型
                data2 = []
                ranking = item.find("span", class_="a-icon-alt")
                if isinstance(ranking, bs4.element.Tag):  # 判断该类元素是否存在，以下的操作也一样
                    ranking = item.find("span", class_="a-icon-alt").get_text().strip()
                    ranking = ranking[0:3]
                date = item.find("span", class_="a-size-base a-color-secondary review-date")
                if isinstance(date, bs4.element.Tag):
                    date = item.find("span", class_="a-size-base a-color-secondary review-date").get_text().strip()
                    date = date[-7::-1]
                    date = date[::-1]
                usefulNum = item.find("span", class_="a-size-base a-color-tertiary cr-vote-text")
                if isinstance(usefulNum, bs4.element.Tag):
                    usefulNum = item.find("span", class_="a-size-base a-color-tertiary cr-vote-text").get_text().strip()
                    usefulNum = usefulNum[-11::-1]
                    usefulNum = usefulNum[::-1]
                comment = item.find("span", class_="a-size-base review-text review-text-content")
                if isinstance(comment, bs4.element.Tag):
                    comment = item.find("span", class_="a-size-base review-text review-text-content").find("span")
                    if isinstance(comment, bs4.element.Tag):
                        comment = item.find("span", class_="a-size-base review-text review-text-content").find(
                            "span").get_text().strip()
                data2.append(ranking)  # 将元素内容添加
                data2.append(date)
                data2.append(usefulNum)
                data2.append(comment)
                time.sleep(0.5)  # 每次处理一条评论休眠0.5秒，意味着每爬一页休眠5秒
                self.dataList.append(data2)  # 将每条评论添加
                print(data2)
            if i % 5 == 0:
                workbook = xlwt.Workbook(encoding="utf-8")  # 打开xls
                worksheet = workbook.add_sheet('sheet1')
                keys = ['ranking', 'date', 'usefulnum', 'comment']  # 设置标题
                for i1, key in enumerate(keys):
                    worksheet.write(0, i1, key)
                for i2, com in enumerate(self.dataList):  # 将信息写入到excel表格
                    for j, item in enumerate(com):
                        worksheet.write(i2 + 1, j, item)
                workbook.save(self.fname)
                print("已保存前" + str(i) + "页")


def main():
    test = Spider(
        "https://www.amazon.com/-/zh/product-reviews/B07F7TLZF4/ref=cm_cr_arp_d_paging_btm_next_2?ie=UTF8&reviewerType=all_reviews&pageNumber=",
        "text.xls")
    test.getData()
    return


main()  # 执行程序
