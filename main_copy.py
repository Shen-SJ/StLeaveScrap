import requests
from requests.cookies import remove_cookie_by_name
from lxml import etree
from urllib import parse
import itertools
import pandas as pd
import copy

import sys

# sys.path.append(r'.')

# 登入用帳密
username = input("請輸入賬號：")
password = input("請輸入密碼：")

# 各種網址
pre_url = r'https://my.ntu.edu.tw/stuLeaveManagement/login.aspx?firstpage=teacher'
login_url = r'https://web2.cc.ntu.edu.tw/p/s/login2/p1.php'
search_url = r'https://my.ntu.edu.tw/stuLeaveManagement/QforTeacher_teacher.aspx'

# 模擬瀏覽器所需的資料內容
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Safari/537.36',
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
    'Referer': r'https://my.ntu.edu.tw/stuLeaveManagement/QforTeacher_teacher.aspx',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'same-origin',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
}

post_headers = {
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

login_data = {'user': username,
              'pass': password,
              'Submit': '登入'
              }

search_data = {
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
}

post_data = {
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
    "ctl00$ContentPlaceHolder1$endDateTextBox":	"2020/02/08",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl02$HiddenField1":	"171700",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl02$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl03$HiddenField1":	"171680",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl03$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl04$HiddenField1":	"171679",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl04$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl05$HiddenField1":	"171654",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl05$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl06$HiddenField1":	"171579",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl06$HiddenField2":	"003+10230",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl07$HiddenField1":	"171577",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl07$HiddenField2":	"003+10230",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl08$HiddenField1":	"171574",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl08$HiddenField2":	"003+10230",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl09$HiddenField1":	"171466",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl09$HiddenField2":	"003+10230",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl10$HiddenField1":	"171431",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl10$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl11$HiddenField1":	"171426",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl11$HiddenField2":	"003+10220",
    "ctl00$ContentPlaceHolder1$GVapprovedList$ctl13$PageDropDownList":	"第2頁"
}

# 開始瀏覽網頁囉
if __name__ == "__main__":
    rs = requests.Session()

    # 先開啟起始網頁取得初始cookies取得權限
    pre_web = rs.get(pre_url,
                     headers=headers,
                     verify=False)
    pre_cookies = requests.utils.dict_from_cookiejar(rs.cookies)
    # 登入學校系統
    login_web = rs.post(login_url,
                        headers=headers,
                        data=login_data)

    login_cookies = requests.utils.dict_from_cookiejar(rs.cookies)
    remove_cookie_by_name(rs.cookies, 'PHPSESSID')  # 不知道 cookie 多了不必要的PHPSESSID會不會怎樣，先刪除好了

    # 直接從這邊要開始查詢了
    # 進入學生請假紀錄系統
    search_web = rs.get(search_url,
                        headers=headers)

    search_cookies = requests.utils.dict_from_cookiejar(rs.cookies)

    # 開始輸入查詢條件查詢了
    ## 要先搞定 aspx 的動態驗證參數，把參數送到查詢表單裡
    selector_eventval = etree.HTML(search_web.text)
    search_data["__EVENTVALIDATION"] = selector_eventval.xpath(r'//*[@id="__EVENTVALIDATION"]/@value')[0]
    search_data["__VIEWSTATE"] = selector_eventval.xpath(r'//*[@id="__VIEWSTATE"]/@value')[0]

    # 因為data內容改變，長度可能改變，要更新表頭資訊
    search_headers['Content-Length'] = str(len(parse.urlencode(search_data)))

    # 得到第一筆資料並儲存成df
    results = rs.post(search_url,
                      headers=search_headers,
                      data=search_data)
    results_cookies = requests.utils.dict_from_cookiejar(rs.cookies)
    df = pd.read_html(results.text)[-1].iloc[0:10, 1:12]

    # 想辦法得到第二頁資料吧
    search_data_copy = copy.copy(search_data)

    search_selector = etree.HTML(results.text)
    search_data_copy["__EVENTVALIDATION"] = search_selector.xpath(r'//*[@id="__EVENTVALIDATION"]/@value')[0]
    search_data_copy["__VIEWSTATE"] = search_selector.xpath(r'//*[@id="__VIEWSTATE"]/@value')[0]
    search_data_copy["__EVENTTARGET"] = "ctl00$ContentPlaceHolder1$GVallLeave$ctl13$lbtnNext"       # 我按"下一頁"所需要送出的參數
    hide_pa = search_selector.xpath(r'//*[contains(@name, "ctl00$ContentPlaceHolder1$GVallLeave$ctl") and contains(@name, "$HiddenField1")]')
    if len(hide_pa) > 10:
        raise Exception('不應該超過10個!!')
    for ele in hide_pa:
        search_data_copy[ele.attrib['name']] = ele.attrib['value']
    search_data_copy['ctl00$ContentPlaceHolder1$GVallLeave$ctl13$PageDropDownList'] = search_selector.xpath(
        r'//*[@name="ctl00$ContentPlaceHolder1$GVallLeave$ctl13$PageDropDownList"]//option[@selected="selected"]/@value')[0]
    del search_data_copy['ctl00$ContentPlaceHolder1$Button1']

    # 因為data內容改變，長度可能改變，要更新表頭資訊
    search_headers['Content-Length'] = str(len(parse.urlencode(search_data_copy)))

    # 得到第二筆資料並儲存成df2
    results2 = rs.post(search_url,
                       headers=search_headers,
                       data=search_data_copy)
    df2 = pd.read_html(results2.text)[-1].iloc[0:10, 1:12]

    pass




