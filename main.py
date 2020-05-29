# 備忘錄
# df[['開始日', '結束日', '簽核時間']].apply(pd.to_datetime)
# 從另外一個課程換到另外一個課程可以不用傳前一個課程的學生隱藏代碼呢

import requests
from requests.cookies import remove_cookie_by_name
from lxml import etree
from urllib import parse
import pandas as pd
import time
import dateutil.parser as prs
import copy
import re
import getpass

# 做個程式開頭好了，不然直接輸入帳密有點詭異...
print(r"+====================================================================+")
print(r"           ********** 學生請假紀錄抓取程式 V2.1.2 **********             ")
print(r"   本程式可以在輸入帳號密碼並確認資訊無誤後,登入學生請假系統,                ")
print(r"   自動抓取教師所有課程於特定搜尋時間內之學生請假紀錄,並在目                 ")
print(r"   前程式所在資料夾匯出成Excel檔案,匯出檔名為:                             ")
print(r"   StudentsLeaveRecord_<起始日期>-<結束日期>.xlsx                       ")
print(r"+--------------------------------------------------------------------+")
print(r"   Author   :SSJ                                                      ")
print(r"   Email    :johnson840205@gmail.com                                  ")
print(r"+====================================================================+")

# 登入用帳密
username = input("請輸入賬號：")
password = getpass.getpass("請輸入密碼：")

# 輸入查詢日期
startdate = prs.parse(input('請輸入查詢起始日期:')).strftime('%Y/%m/%d')     # 網頁伺服器只接受 YYY/MM/DD
enddate = prs.parse(input('請輸入查詢結束日期:')).strftime('%Y/%m/%d')       # 網頁伺服器只接受 YYY/MM/DD


class StLeaveScrap:
    # 網址
    pre_url = r'https://my.ntu.edu.tw/stuLeaveManagement/login.aspx?firstpage=teacher'
    login_url = r'https://web2.cc.ntu.edu.tw/p/s/login2/p1.php'
    search_url = r'https://my.ntu.edu.tw/stuLeaveManagement/QforTeacher_teacher.aspx'
    search_detail_url = r'https://my.ntu.edu.tw/stuLeaveManagement/LeaveDetail_teacher.aspx'

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
        "__EVENTTARGET": "",
        "__EVENTARGUMENT": "",
        "__VIEWSTATE": "/wEPDwULLTE4MDkxNjYwMjMPZBYCZg9kFgICAw9kFgICAQ9kFgwCAw8PZA8QFgFmFgEWAh4OUGFyYW1ldGVyVmFsdWVkFgECA2RkAgUPD2QPDxQrAAEWBh4ETmFtZQUIdGVhX2NvZGUeDERlZmF1bHRWYWx1ZQUGMDAzMDQ1HwBkFCsBAQIDZGQCBw8PZA8QFgFmFgEWAh8ABQUyNDY0MhYBAgVkZAIJDxAPFgQeC18hRGF0YUJvdW5kZx4HVmlzaWJsZWhkEBUBBjAwMzA0NRUBBjAwMzA0NRQrAwFnFgBkAg0PEA8WAh8DZ2QQFQJ65YWo5rCR5ZyL6Ziy5pWZ6IKy6LuN5LqL6KiT57e06Kqy56iL77yN5ZyL6Zqb5oOF5YuiIEFsbC1vdXQgRGVmZW5zZSBFZHVjYXRpb24gTWlsaXRhcnkgVHJhaW5pbmcgLSBJbnRlcm5hdGlvbmFsIFNpdHVhdGlvbnNz5YWo5rCR5ZyL6Ziy5pWZ6IKy6LuN5LqL6KiT57e06Kqy56iL77yN5ZyL6Ziy56eR5oqAIEFsbC1vdXQgRGVmZW5zZSBFZHVjYXRpb24gTWlsaXRhciBUcmFpbmluZyAtIERlZmVuc2UgVGVjaG5vbG9neRUCCTAwMyAxMDIyMAkwMDMgMTAyMzAUKwMCZ2dkZAIVDxAPFgIfA2dkEBUHBuWFqOmDqAvnl4XlgYcgc2ljaw/kuovlgYcgcGVyc29uYWwP5YWs5YGHIG9mZmljaWFsEOeUouWBhyBtYXRlcm5pdHkT55Sf55CG5YGHIG1lbnN0cnVhbA7llqrlgYcgZnVuZXJhbBUHAAEyATMBNAE1AjExATcUKwMHZ2dnZ2dnZ2RkGAEFJGN0bDAwJENvbnRlbnRQbGFjZUhvbGRlcjEkR1ZhbGxMZWF2ZQ88KwAKAQhmZGwahZ/QT02yJ69DM4ZODhYDYm63",
        "__VIEWSTATEGENERATOR": "1C8E3960",
        "__EVENTVALIDATION": "/wEWHQKy88f5DAKy3qeGBwKy3pPdDwLRhuutCwLB6cHDBwLe6cHDBwLi0oa2DgLPurjaDwLB1ZK0AwLC1ZK0AwLD1ZK0AwLE1ZK0AwLA1d63AwLG1ZK0AwLwkuyECgLorf/gBQLQ3ZzYBQL2yrfJAQL8+Mu7DwKMt5/YBQKNuNfdBgLslbasCALgx48/AonJgv0PAo7Tk/8JAtmq8MUGAsPP7NsEAorA9PYHAoDiyWPNr+i1K3LavI0cS3R5YMrfU/OS5Q==",
        "ctl00$ContentPlaceHolder1$DDLcourse": "003 10220",
        "ctl00$ContentPlaceHolder1$DDLisAbroad": "",
        "ctl00$ContentPlaceHolder1$nameTextBox": "",
        "ctl00$ContentPlaceHolder1$DDLLeaveType": "",
        "ctl00$ContentPlaceHolder1$DDLstatus": "",
        "ctl00$ContentPlaceHolder1$DDLformType": "",
        "ctl00$ContentPlaceHolder1$startDateTextBox": r"2020/03/23",
        "ctl00$ContentPlaceHolder1$endDateTextBox": r"2020/05/23",
        "ctl00$ContentPlaceHolder1$Button1": "查詢"
    }       # 可以送出第一次搜尋用

    # 會用到的參數
    courses_dict = None  # 有哪些課要抓
    search_web = None  # QforTeacher 的網頁內容，會隨之更改
    search_web_etree = None  # search_web 的 etree 版本， xpath 用

    # 我抓到的資料
    st_leave_data = None

    def __init__(self, account, pas, startdate=None, enddate=None):
        """
        一個可以抓取學生請假紀錄的類別，包含了網頁 post 方法、資料抓取方法、資料過濾方法等。

        :param account: 登入的帳號
        :param pas: 登入的密碼
        :param startdate: 資料的起始日期
        :param enddate: 資料的結束日期
        """
        # 登入用的帳密字典
        self.login_data = {'user': account,
                           'pass': pas,
                           'Submit': '登入'
                           }

        # 建立一個爬蟲的物件
        self.rs = requests.Session()

        # 輸入起始日期與結束日期
        self.startdate = startdate
        self.enddate = enddate

    def login(self):
        """
        輸入帳號密碼登入學生請假紀錄系統。

        :return: 登入後的網頁 respond 物件。
        """
        # 先開啟起始網頁取得初始cookies取得權限
        pre_web = self.rs.get(self.pre_url,
                              headers=self.nor_headers,
                              verify=False)     # 用 fiddler 驗證不會過，所以要關閉
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
        取得需要抓取的所有課程名稱與其編號

        :param web_text: 資料來源網頁內容
        :return: 有課程資訊與其值的字典
        """
        courses_dict = {}
        selector_eventval = etree.HTML(web_text)
        data = selector_eventval.xpath(r'//*[@name="ctl00$ContentPlaceHolder1$DDLcourse"]//option')
        for ele in data:
            courses_dict[ele.attrib['value']] = ele.text

        return courses_dict

    def change_aspnet_arg(self):
        """
        變更三個必要的 asp.net 隱藏參數

        :return: None
        """
        self.search_post_data["__VIEWSTATE"] = self.search_web_etree.xpath(r'//*[@id="__VIEWSTATE"]/@value')[0]
        self.search_post_data["__EVENTVALIDATION"] = \
            self.search_web_etree.xpath(r'//*[@id="__EVENTVALIDATION"]/@value')[0]
        self.search_post_data["__VIEWSTATEGENERATOR"] = \
            self.search_web_etree.xpath(r'//*[@id="__VIEWSTATEGENERATOR"]/@value')[0]  # 其實目前他不會變，但怕有萬一還是從頁面抓來

    def change_hidden_field_pa(self):
        """
        添加或修改每位學生具有的 ctl00$ContentPlaceHolder1$GVallLeave$ctl[num]$HiddenField1，一個頁面大於10組

        :return: None
        """
        hide_pa = etree.HTML(self.search_web.text).xpath(
            r'//*[contains(@name, "ctl00$ContentPlaceHolder1$GVallLeave$ctl") and contains(@name, "$HiddenField1")]')  # 取得隱藏參數
        if len(hide_pa) > 10:
            raise Exception('不應該超過10個!!')
        for ele in hide_pa:  # 隱藏的那十個參數
            self.search_post_data[ele.attrib['name']] = ele.attrib['value']

    def get_course_data(self, course_id):
        """
        取得指定課程的學生請假紀錄

        :param course_id: 指定課程的代碼
        :return: 學生請假紀錄 dataframe
        """
        # TODO: 這段感覺超級亂，之後有機會再改
        # 把課程代碼 key-in 到 search_post_data
        self.search_post_data["ctl00$ContentPlaceHolder1$DDLcourse"] = course_id
        self.search_post_data["ctl00$ContentPlaceHolder1$startDateTextBox"] = self.startdate
        self.search_post_data["ctl00$ContentPlaceHolder1$endDateTextBox"] = self.enddate

        ## 先取得第一頁吧
        # 更改 search_headers
        self.search_headers['Content-Length'] = str(len(parse.urlencode(self.search_post_data)))

        # 把 search_post_data 備份，之後的變更不想應用到類別屬性，才能確保都是在 QforTeacher 的 root_web 搜尋不同課程
        search_post_data_copy = copy.copy(self.search_post_data)

        # 取得第一頁並抓取資料
        time.sleep(1)  # 怕抓太快對伺服器造成負擔
        self.search_web = self.rs.post(self.search_url,
                                       headers=self.search_headers,
                                       data=self.search_post_data)
        self.search_web_etree = etree.HTML(self.search_web.text)

        # 因為第一頁如果<=10，就不會有頁數選單-->df擷取的位置不能少，本來 ".iloc[0:-1" 是因為最後一row為表單所以要刪除
        if int(self.search_web_etree.xpath(r'//*[@id="ctl00_ContentPlaceHolder1_recordCountLabel"]')[0].text) <= 10:
            data = pd.read_html(self.search_web.text)[-1].iloc[0:, 1:12]
        else:
            data = pd.read_html(self.search_web.text)[-1].iloc[0:-1, 1:12]

        # TODO: 進一步確認"簽核中"的狀況

        # 算算這次你要抓幾頁，迴圈要跑幾次
        total_pages = len(self.search_web_etree.xpath(
            '//*[@name="ctl00$ContentPlaceHolder1$GVallLeave$ctl13$PageDropDownList"]//option'))

        ## 取得第二頁之後的資料吧
        for _ in range(0, total_pages - 1):
            # 更改 search_post_data
            self.change_aspnet_arg()  # 更改 asp.net 三個必要參數
            self.search_post_data["__EVENTTARGET"] = "ctl00$ContentPlaceHolder1$GVallLeave$ctl13$lbtnNext"  # 我按"下一頁"所需要送出的參數
            self.change_hidden_field_pa()   # 更改或添加 HiddenField 參數(每個人都有一個，不超過10個)
            self.search_post_data['ctl00$ContentPlaceHolder1$GVallLeave$ctl13$PageDropDownList'] = \
                etree.HTML(self.search_web.text).xpath(
                    r'//*[contains(@name, "ctl00$ContentPlaceHolder1$GVallLeave$ctl") and contains(@name, "$PageDropDownList")]//option[@selected="selected"]/@value')[0]  # 選中的頁面，通常最後一頁不會跑到這，那 ctl13 應該就沒問題吧
            if 'ctl00$ContentPlaceHolder1$Button1' in self.search_post_data.keys():
                del self.search_post_data['ctl00$ContentPlaceHolder1$Button1']  # 第一筆資料要按"查詢"才取得，但第二筆開始不用，所以要刪掉

            # 更改 search_headers
            self.search_headers['Content-Length'] = str(len(parse.urlencode(self.search_post_data)))

            # 得到第下筆資料
            time.sleep(1)  # 怕抓太快對伺服器造成負擔
            self.search_web = self.rs.post(self.search_url,
                                           headers=self.search_headers,
                                           data=self.search_post_data)
            self.search_web_etree = etree.HTML(self.search_web.text)
            # 萃取出所需的資料後變成 df
            datan = pd.read_html(self.search_web.text)[-1].iloc[0:-1, 1:12]
            # TODO: 進一步確認"簽核中"的狀態，這邊就要抓到的網路頁面具有的隱藏參數了
            data = pd.concat([data, datan],
                             axis='index',
                             ignore_index=True)

        # 還原 search_post_data
        self.search_post_data = search_post_data_copy
        # 因為我這邊的概念很像是平行查詢，每個課程的查詢都是從最初、還沒有送出 post 前的那個 QforTeacher 頁面的 cookies 開始查詢，
        # 如下圖:
        #
        #                       * : login to SignList_teacher
        #                       |
        #                       * : goto QforTeacher
        #                      / \
        #                     /   \
        #     search 10230 : *     * : search 10220
        #                    |     |
        #  get the results : *     * : get the results
        #

        return data

    def export2excel(self):
        """
        將抓到的 st_leave_data 輸出成 excel. 不同課程會分別以不同 sheet 隔開。

        :return: None.
        """
        # 我的檔名是啥
        filename = f"StudentsLeaveRecord_{self.startdate.replace(r'/', '')}-{self.enddate.replace(r'/', '')}"

        # pandas 要匯出多 sheets excel 要先創建一個 ExcelWriter 物件，然後在這樣寫進去
        with pd.ExcelWriter(f'{filename}.xlsx') as writer:
            for course_name in self.st_leave_data.keys():
                # 因為 course_name 太長了，所以想要變短一些，又發現該名稱包含中文與英文翻譯，打算把英文翻譯去掉達到變短的效果
                course_translation = re.search(r'[a-zA-Z\s-]{2,}', course_name).group()
                course_name_short = course_name.replace(course_translation, '')
                # 把指定課程資料存在縮短過的 course_name sheet
                self.st_leave_data[course_name].to_excel(writer,
                                                         sheet_name=course_name_short,
                                                         index=False)

    def scrapping(self):
        """
        抓取該帳號內所有課程指定日期間學生請假紀錄。

        :return: a dataframe.
        """
        # 先登入學生請假系統
        self.login()  # 已取得登入cookies

        # 進入 QforTeacher 頁面並存起來
        self.search_web = self.rs.get(self.search_url,
                                      headers=self.nor_headers)
        self.search_web_etree = etree.HTML(self.search_web.text)

        # 更改 search_post_data
        self.change_aspnet_arg()  # 更改 asp.net 三個必要參數

        # 更改 search_headers
        self.search_headers['Content-Length'] = str(len(parse.urlencode(self.search_post_data)))

        # 先確定有幾個課程要抓以及取得課程名稱與代碼
        self.courses_dict = self.get_course_dict(self.search_web.text)

        # 開始抓吧!!
        self.st_leave_data = {}
        for course_id in self.courses_dict.keys():
            self.st_leave_data[self.courses_dict[course_id]] = self.get_course_data(course_id)

        return self.st_leave_data


class LoginError(Exception):
    """登入失敗用"""
    pass


# 開始瀏覽網頁囉
if __name__ == "__main__":
    s = StLeaveScrap(username, password, startdate=startdate, enddate=enddate)
    df = s.scrapping()
    s.export2excel()

    pass
