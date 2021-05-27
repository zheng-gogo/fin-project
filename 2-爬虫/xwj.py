# 编写人员：梁雨
# 编写时间：2021/1/29 11:45
# 函数功能：
import datetime
import requests
from bs4 import BeautifulSoup
from urllib.error import HTTPError, URLError
import openpyxl

total = 10000

if __name__ == '__main__':
    date = datetime.date.today()
    urls = 'http://stock.jrj.com.cn/xwk/'
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/88.0.4324.96 Safari/537.36 Edg/88.0.705.53 '
    }
    '''http://stock.jrj.com.cn/xwk/202101/20210129_2.shtml'''
    path = '../../Table/00_other/xwj.xlsx'
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    count = 1
    url_set = set()
    tag = True
    info = ['', '', '', '']
    while tag:
        urlx = urls + date.__format__("%Y%m") + '/' + date.__format__("%Y%m%d")
        date = date + datetime.timedelta(days=-1)
        pages = 20
        i = 1
        if urlx in url_set:
            continue
        else:
            url_set.add(urlx)
        while i <= pages and tag:
            try:
                url = urlx + '_' + str(i) + '.shtml'
                page_text = requests.get(url=url, headers=headers).text
                soup = BeautifulSoup(page_text, 'lxml')
                a_list = soup.select('.list > li > a')
                url_list = []
                if i == 1:
                    pages = len(soup.select('.page_newslib > a')) - 3
                    if pages <= 0:
                        break
                for item in a_list:
                    url_list.append(item['href'])
                for item in url_list:
                    if item in url_set:
                        continue
                    else:
                        url_set.add(item)
                    print(count, item)
                    news_text = requests.get(url=item, headers=headers).text
                    news_soup = BeautifulSoup(news_text, 'lxml')
                    info[1] = news_soup.find('p', class_='inftop').span.text.strip()
                    info[2] = news_soup.h1.text.strip()
                    info[3] = news_soup.find('div', class_='texttit_m1').text.replace(' ', '').replace('\n', '').strip()
                    for i in range(1, 4):
                        sheet.cell(count, i, info[i])
                    count += 1
                    if count > total:
                        tag = False
                        break
                i += 1
            except URLError as e:
                print(e)
                break
            except HTTPError as e:
                print(e)
                break
            except Exception as e:
                print(e)
                pages = 0
            finally:
                workbook.save(path)
    workbook.save(path)
