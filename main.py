# 備忘錄
# df[['開始日', '結束日', '簽核時間']].apply(pd.to_datetime)

import requests
from requests.cookies import remove_cookie_by_name
from lxml import etree
from urllib import parse
import itertools
import pandas as pd
import numpy as np
import time
import dateutil.parser as prs

# 登入用帳密
username = input("請輸入賬號：")
password = input("請輸入密碼：")

# 輸入查詢日期
startdate = pd.Timestamp((prs.parse(input('請輸入查詢起始日期:')).date()))
enddate = pd.Timestamp((prs.parse(input('請輸入查詢結束日期:')).date())) + pd.Timedelta(days=1)
# TODO: 任意格式的時間轉換成特定時間


class StLeaveScrap:
    # 網址
    pre_url = r'https://my.ntu.edu.tw/stuLeaveManagement/login.aspx?firstpage=teacher'
    login_url = r'https://web2.cc.ntu.edu.tw/p/s/login2/p1.php'
    search_url = r'https://my.ntu.edu.tw/stuLeaveManagement/QforTeacher_teacher.aspx'

    # 表頭
    nor_headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:72.0) Gecko/20100101 Firefox/72.0',
                   'Connection': 'keep-alive'
                   }
    search_headers = {
        'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36',
        'Accept': r'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Accept-Encoding': 'gzip, deflate, br',
        'Accept-Language': 'zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7,zh-CN;q=0.6,it;q=0.5,ja;q=0.4',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Content-Length': '2097',
        'Content-Type': 'application/x-www-form-urlencoded',
        'Host': 'my.ntu.edu.tw',
        'Origin': r'https://my.ntu.edu.tw',
        'Referer': r'https://my.ntu.edu.tw/stuLeaveManagement/SignList_teacher.aspx',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
    }

    # 表單
    search_post_data = {
        "ctl00_ContentPlaceHolder1_ScriptManager1_HiddenField": ";;AjaxControlToolkit, Version=1.0.10618.0, Culture=neutral, PublicKeyToken=28f01b0e84b6d53e:zh-TW:bc82895f-eb24-48f8-a8ba-a354eb9c74da:e2e86ef9:a9a7729d:9ea3f0e2:9e8e87e9:1df13a87:4c9865be:ba594826:507fcf1b:c7a4182e",
        "__EVENTTARGET": "ctl00$ContentPlaceHolder1$GVapprovedList$ctl13$lbtnNext",
        "__EVENTARGUMENT": "",
        "__LASTFOCUS": "",
        "__VIEWSTATE": "/wEPDwULLTE4MDkxNjYwMjMPZBYCZg9kFgICAw9kFgICAQ9kFgwCAw8PZA8QFgFmFgEWAh4OUGFyYW1ldGVyVmFsdWVkFgECA2RkAgUPD2QPDxQrAAEWBh4ETmFtZQUIdGVhX2NvZGUeDERlZmF1bHRWYWx1ZQUGMDAzMDQ1HwBkFCsBAQIDZGQCBw8PZA8QFgFmFgEWAh8ABQUyNDY0MhYBAgVkZAIJDxAPFgQeC18hRGF0YUJvdW5kZx4HVmlzaWJsZWhkEBUBBjAwMzA0NRUBBjAwMzA0NRQrAwFnFgBkAg0PEA8WAh8DZ2QQFQJ65YWo5rCR5ZyL6Ziy5pWZ6IKy6LuN5LqL6KiT57e06Kqy56iL77yN5ZyL6Zqb5oOF5YuiIEFsbC1vdXQgRGVmZW5zZSBFZHVjYXRpb24gTWlsaXRhcnkgVHJhaW5pbmcgLSBJbnRlcm5hdGlvbmFsIFNpdHVhdGlvbnNz5YWo5rCR5ZyL6Ziy5pWZ6IKy6LuN5LqL6KiT57e06Kqy56iL77yN5ZyL6Ziy56eR5oqAIEFsbC1vdXQgRGVmZW5zZSBFZHVjYXRpb24gTWlsaXRhciBUcmFpbmluZyAtIERlZmVuc2UgVGVjaG5vbG9neRUCCTAwMyAxMDIyMAkwMDMgMTAyMzAUKwMCZ2dkZAIVDxAPFgIfA2dkEBUHBuWFqOmDqAvnl4XlgYcgc2ljaw/kuovlgYcgcGVyc29uYWwP5YWs5YGHIG9mZmljaWFsEOeUouWBhyBtYXRlcm5pdHkT55Sf55CG5YGHIG1lbnN0cnVhbA7llqrlgYcgZnVuZXJhbBUHAAEyATMBNAE1AjExATcUKwMHZ2dnZ2dnZ2RkGAEFJGN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkR1ZhbGxMZWF2ZQ88KwAKAQhmZGwahZ/QT02yJ69DM4ZODhYDYm63",
        "__VIEWSTATEGENERATOR": "0C3B263D",
        "__SCROLLPOSITIONX": "0",
        "__SCROLLPOSITIONY": "213",
        "__EVENTVALIDATION": "/wEWHQKy88f5DAKy3qeGBwKy3pPdDwLRhuutCwLB6cHDBwLe6cHDBwLi0oa2DgLPurjaDwLB1ZK0AwLC1ZK0AwLD1ZK0AwLE1ZK0AwLA1d63AwLG1ZK0AwLwkuyECgLorf/gBQLQ3ZzYBQL2yrfJAQL8+Mu7DwKMt5/YBQKNuNfdBgLslbasCALgx48/AonJgv0PAo7Tk/8JAtmq8MUGAsPP7NsEAorA9PYHAoDiyWPNr+i1K3LavI0cS3R5YMrfU/OS5Q==",
        "ctl00$ContentPlaceHolder1$DDLapplylist": "all",
        "ctl00$ContentPlaceHolder1$commentTextBox": "",
        "ctl00$ContentPlaceHolder1$startDateTextBox": "2020/02/08",
        "ctl00$ContentPlaceHolder1$endDateTextBox": "2020/02/08",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl02$HiddenField1": "171700",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl02$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl03$HiddenField1": "171680",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl03$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl04$HiddenField1": "171679",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl04$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl05$HiddenField1": "171654",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl05$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl06$HiddenField1": "171579",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl06$HiddenField2": "003+10230",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl07$HiddenField1": "171577",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl07$HiddenField2": "003+10230",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl08$HiddenField1": "171574",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl08$HiddenField2": "003+10230",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl09$HiddenField1": "171466",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl09$HiddenField2": "003+10230",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl10$HiddenField1": "171431",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl10$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl11$HiddenField1": "171426",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl11$HiddenField2": "003+10220",
        "ctl00$ContentPlaceHolder1$GVapprovedList$ctl13$PageDropDownList": "第2頁"
    }

    # 會用到的參數
    courses_dict = None      # 有哪些課要抓

    def __init__(self, account, pas, startdate=None, enddate=None): # TODO: 可以指定開始日與結束日
        # 登入用的帳密字典
        self.login_data = {'user': account,
                           'pass': pas,
                           'Submit': '登入'
                           }

        # 建立一個爬蟲的物件
        self.rs = requests.Session()

        self.startdate = startdate
        self.enddate = enddate

    def get_search_post_data(self, web_text):
        ## 搞定 aspx 的三個動態驗證參數，把參數送到查詢表單裡
        selector_eventval = etree.HTML(web_text)
        self.search_post_data["__VIEWSTATE"] = selector_eventval.xpath(r'//*[@id="__VIEWSTATE"]/@value')[0]
        self.search_post_data["__VIEWSTATEGENERATOR"] = selector_eventval.xpath(r'//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]
        self.search_post_data["__SCROLLPOSITIONX"] = selector_eventval.xpath(r'//*[@id="__SCROLLPOSITIONX"]/@value')[0]
        self.search_post_data["__SCROLLPOSITIONY"] = selector_eventval.xpath(r'//*[@id="__SCROLLPOSITIONY"]/@value')[0]
        self.search_post_data["__EVENTVALIDATION"] = selector_eventval.xpath(r'//*[@id="__EVENTVALIDATION"]/@value')[0]

        self.search_post_data["ctl00$ContentPlaceHolder1$startDateTextBox"] = \
        selector_eventval.xpath(r'//*[@id="ctl00_ContentPlaceHolder1_startDateTextBox"]/@value')[0]
        self.search_post_data["ctl00$ContentPlaceHolder1$endDateTextBox"] = \
        selector_eventval.xpath(r'//*[@id="ctl00_ContentPlaceHolder1_endDateTextBox"]/@value')[0]
        # 頁面中的隱藏參數，一筆資料會有一個隱藏參數，一頁會有十筆資料，因此會有10個隱藏參數
        for i, j in itertools.product([str(num).zfill(2) for num in range(2, 12)], ['1', '2']):
            self.search_post_data[f"ctl00$ContentPlaceHolder1$GVapprovedList$ctl{i}$HiddenField{j}"] = selector_eventval.xpath(
                rf'//*[@id="ctl00_ContentPlaceHolder1_GVapprovedList_ctl{i}_HiddenField{j}"]/@value')[0]
        # 連你那頁的頁面選擇單的值也要傳過去...
        self.search_post_data["ctl00$ContentPlaceHolder1$GVapprovedList$ctl13$PageDropDownList"] = selector_eventval.xpath(
            r'//*[@id="ctl00_ContentPlaceHolder1_GVapprovedList_ctl13_PageDropDownList"]//option[@selected="selected"]/@value')[0]

        return self.search_post_data

    def get_post_heads(self):
        # 因為data內容改變，長度可能改變，要更新表頭資訊
        self.search_headers['Content-Length'] = str(len(parse.urlencode(self.search_post_data)))

        return self.search_headers

    def login(self):
        """
        登入學生請假系統

        :return: HTML for login web.
        """
        # 先開啟起始網頁取得初始cookies取得權限
        pre_web = self.rs.get(self.pre_url,
                              headers=self.nor_headers)
        pre_cookies = requests.utils.dict_from_cookiejar(self.rs.cookies)

        # 登入學校系統
        login_web = self.rs.post(self.login_url,
                                 headers=self.nor_headers,
                                 data=self.login_data)
        login_cookies = requests.utils.dict_from_cookiejar(self.rs.cookies)
        remove_cookie_by_name(self.rs.cookies, 'PHPSESSID')  # 不知道 cookie 多了不必要的PHPSESSID會不會怎樣，先刪除好了

        # 檢查登入有沒有失敗
        if login_web.url != r'https://my.ntu.edu.tw/stuLeaveManagement/SignList_teacher.aspx':
            raise LoginError('登入失敗!!!')

        return login_web

    def get_course_dict(self, web_text):
        """
        取得需要抓取地所以課程名稱與其編號

        :param web_text: 資料來源網頁內容
        :return: 有課程資訊與其值的字典
        """
        courses_dict = {}
        selector_eventval = etree.HTML(web_text)
        data = selector_eventval.xpath(r'//*[@name="ctl00$ContentPlaceHolder1$DDLcourse"]//option')
        for ele in data:
            courses_dict[ele.text] = ele.attrib['value']

        return courses_dict

    def scrap_data(self, web_text):
        """
        將 web_text 中所需表格抓下來變成 dataframe

        :param web_text:
        :return:
        """
        selector_eventval = etree.HTML(web_text)
        data = selector_eventval.xpath(rf'//*[@id="ctl00_ContentPlaceHolder1_GVapprovedList"]/tr/td[position()<last() and position()>1]/text()')
        data = np.asarray(data).reshape((-1, 8))
        df = pd.DataFrame(data, columns=self.get_table_head(web_text))
        return df

    def filter_by_date(self, df):
        df = df[df['簽核時間'] < self.enddate]
        df = df[df['簽核時間'] > self.startdate]
        return df

    def filter_by_confirm(self, df):
        df = df[df['我的簽核'] == '核准']
        return df

    def get_table_head(self, web_text):
        selector_eventval = etree.HTML(web_text)
        # 先取得欄位名稱
        table_head = selector_eventval.xpath(rf'//*[@id="ctl00_ContentPlaceHolder1_GVapprovedList"]/tr/th[position()>1 and position()<last()]/text()')
        return table_head

    def scrapping(self):
        # 先登入學生請假系統
        self.login()    # 已取得登入cookies

        # 進入 QforTeacher 頁面
        search_web = self.rs.get(self.search_url,
                                 headers=self.nor_headers)

        # 先確定有幾個課程要抓
        self.courses_dict = self.get_course_dict(search_web.text)

        return

        # # 下面就是開始抓資料了
        # df = self.scrap_data(login_web.text)
        # df[['開始日', '結束日', '簽核時間']] = df[['開始日', '結束日', '簽核時間']].apply(pd.to_datetime)
        #
        # df = self.filter_by_date(df)    # 根據搜尋日期區間篩選資料
        #
        # # 看看第二頁的內容 ((第一頁在上面登錄完後出現
        # self.get_search_post_data(login_web.text)
        # self.get_post_heads()
        #
        # while True:
        #     origin_size = df.size
        #     # 取得更新後的頁面
        #     time.sleep(3)
        #     results = self.rs.post(self.search_url,
        #                            headers=self.search_headers,
        #                            data=self.search_post_data)
        #     results_cookies = requests.utils.dict_from_cookiejar(self.rs.cookies)
        #
        #     # 萃取出所需的資料後變成 df
        #     df = pd.concat([df, self.scrap_data(results.text)], axis='index', ignore_index=True)
        #     df[['開始日', '結束日', '簽核時間']] = df[['開始日', '結束日', '簽核時間']].apply(pd.to_datetime)
        #
        #     df = self.filter_by_date(df)    # 根據搜尋日期區間篩選資料
        #
        #     # 如果增加資料後的大小跟沒增加前一樣(但不是沒資料)，代表資料以抓完畢
        #     if df.size == origin_size and df.size != 0:
        #         break
        #
        #     # 更新取得下頁資料所需的參數
        #     self.get_search_post_data(results.text)
        #     self.get_post_heads()
        #
        # df = self.filter_by_confirm(df)     # 只保存核准過的資料
        #
        # return results, results_cookies, df


class LoginError(Exception):
    """登入失敗用"""
    pass


# 開始瀏覽網頁囉
if __name__ == "__main__":
    s = StLeaveScrap(username, password, startdate=startdate, enddate=enddate)
    # r, r_cookie, df = s.scrapping()
    s.scrapping()

    pass
