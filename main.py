# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
import requests
import urllib3
from lxml import etree
import pandas as pd
from time import sleep
import keyboard
from pandas import DataFrame
import undetected_chromedriver as uc
import os
import re
import configparser
import logging
import logging.handlers

config = configparser.ConfigParser()
config.read("my.ini", encoding="utf-8")
proxy_host = config.get("proxy", "host")
sleep_time = config.get("work", "sleep_time")
file_name = config.get("work", "file_name")
file_output_name = config.get("work", "file_output_name")
one_open_browser_count = config.get("work", "one_open_browser_count")
second_search_source_file_name = config.get("work", "second_search_source_file_name")
second_search_file_output_name = config.get("work", "second_search_file_output_name")
url_home = 'https://www.cyberbackgroundchecks.com'
cookie = ''
user_agent = ''
procedure_file_name = './procedure/'
result_file_name = './result/'


def open_browser_get_cookie(url):
    driver = uc.Chrome()
    driver.get(url)
    sleep(2)
    cookies = driver.get_cookies()
    user_agent_ = driver.execute_script('return navigator.userAgent')
    cookie_result = ''
    count = 0
    for cookie1 in cookies:
        name = cookie1.get('name')
        value = cookie1.get('value')
        if count == (len(cookies) - 1):
            cookie_result = cookie_result + name + '=' + value
        else:
            cookie_result = cookie_result + name + '=' + value + '; '
        count = count + 1
    global cookie
    cookie = cookie_result
    global user_agent
    user_agent = user_agent_
    driver.quit()


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def get_cookie():
    session = requests.session()
    session.get(url=url_home)


def files_mergers():
    df_data = pd.read_excel(io='./result_file_0.xlsx')
    df_data1 = pd.read_excel(io='./result_file_1.xlsx')
    df_data2 = pd.read_excel(io='./result_file_2.xlsx')
    concat = pd.concat([df_data, df_data1, df_data2])
    pd.concat()
    concat.to_excel('all_data.xlsx')


def split_file():
    df_data = pd.read_excel(io='./4-20(2).xlsx')
    i = len(df_data)
    len_ = int(i / 500)
    i_ = i % 500
    if i_ > 0:
        len_ = len_ + 1
    for num in range(len_):
        if num == len_:
            df_data1 = df_data.iloc[num * 500:len_, :]
        else:
            df_data1 = df_data.iloc[num * 500:(num + 1) * 500, :]
        file_name = 'file_' + str(num) + '.xlsx'
        print('生成excel' + str(num))
        df_data1.to_excel(file_name)


# See PyCharm help at https://www.jetbrains.com/help/pycharm/
def get_net_data(filename):
    with urllib3.ProxyManager(proxy_host) as http:
        df_data = pd.read_excel(io='./' + filename)
        columns_ = df_data.columns.tolist()
        columns_.insert(16, 'pid')
        columns_.insert(17, '匹配方式')
        columns_.insert(18, '匹配等级')
        columns_.insert(18, '手机号类型')
        df_data.reindex(columns=columns_)
        for i in range(len(df_data)):
            print(str(i))
            data1 = trim(str(df_data.loc[i, "买家姓名"]).replace(u'\xa0', ' '))
            data1 = re.sub(r"\d+", '', data1).split('/')[0]
            data2 = trim(str(df_data.loc[i, "县/城市"]).replace(u'\xa0', ' ')).split('/')[0]
            data3 = trim(str(df_data.loc[i, "州/地区"]).replace(u'\xa0', ' ')).split('/')[0]
            data4 = trim(str(df_data.loc[i, "收件人"]).replace(u'\xa0', ' '))
            data4 = re.sub(r"\d+", '', data4).split('/')[0]
            data5 = trim(str(df_data.loc[i, "地址"]).replace(u'\xa0', ' ')).split('/')[0]
            df_data.loc[i, "pid"] = 'None'
            if data1 is None or data1 == '' or data1 == 'nan' or data1 == 'None' or not isinstance(data1, str) \
                    or not isinstance(data3, str):
                df_data.loc[i, "pid"] = 'None'
                df_data.loc[i, "匹配等级"] = 'None'
                continue
            door_plate_num = -1
            if data5 is not None and isinstance(data5, str):
                d = data5.split(' ')[0]
                if d.isdigit():
                    door_plate_num = int(d)
            temp1 = data1.lower().replace('.', '').replace(',', '')
            temp4 = data4.lower().replace('.', '').replace(',', '')
            search_name = temp1

            if data4 is not None:
                if temp1 in temp4:
                    search_name = temp4
                    if temp1 == temp4:
                        df_data.loc[i, "匹配方式"] = '买家姓名'
                    else:
                        df_data.loc[i, "匹配方式"] = '收件人'
                else:
                    df_data.loc[i, "匹配方式"] = '买家姓名'

            if len(temp1.split(' ')[0]) == 1:
                search_name = temp4
                df_data.loc[i, "匹配方式"] = '收件人'

            address_pp = False
            if len(temp1.split(' ')) == 1 and search_name == temp1:
                address_pp = True

            replace = re.sub(u"\\(.*?\\)", " ", re.sub(u"\\s\\(.*?\\)\\s", " ", search_name)).replace('  ',
                                                                                                      ' ').replace(' ',
                                                                                                                   '-').lower()
            # print(replace + '\t' + data2.lower().replace(' ', '-') + '\t' + data3.lower())
            if address_pp:
                df_data.loc[i, "匹配方式"] = '地址'
                url = url_home + '/address/' + data5.replace('  ', ' ').replace(' ', '-').lower() + '/' + data2.lower() \
                    .replace(' ', '-') + '/' + data3.lower().replace(' ', '-')
            else:
                url = url_home + '/people/' + replace + '/' + data3.lower() + '/' + data2.lower().replace(' ', '-')

            print(url)
            if i % int(one_open_browser_count) == 0:
                open_browser_get_cookie(url)
            if i == 2:
                a = [2, 3]
                b = a[3]
            headers = {
                'User-Agent': user_agent,
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-',
                'Accept-Language': 'zh-CN,zh;q=0.9',
                'Connection': 'keep-alive',
                'Cookie': cookie,
            }
            response = http.request('GET', url, headers=headers)

            if response.status == 403:
                print(cookie)
                print(headers)
                df_data.to_excel('zd_pid_result.xlsx')
                print('出现' + str(response.status) + '错误，任务中断')
                break

            if response.status != 200:
                print('出现' + str(response.status) + '错误')
                df_data.loc[i, "pid"] = '404'
                df_data.loc[i, "匹配等级"] = '404'
                continue

            tree = etree.HTML(response.data)
            # result = tree.xpath('/html/body/div[3]/div[3]/div/div[2]/div/div[2]/div[7]/div/span')
            # result2 = tree.xpath("//div[starts-with(@id,'opt-out-disabled-')]")
            # for it in result:
            #     print(etree.tostring(it, encoding='utf-8').decode('utf-8'))
            table = tree.xpath("//div[@class='card card-hover']")
            address_dpn_res = []
            address_cj_dpn_res = []
            address_ok_res = []
            address_cj_res = []
            address_all_res = []
            res_names = []
            res_names_cj = []
            for it in table:
                # address = it.xpath("//p[contains(@class,'address-current')]")
                it_tree = etree.tostring(it, encoding='utf-8').decode('utf-8').lower()
                result_final = etree.HTML(it_tree).xpath("//div[starts-with(@id,'opt-out-disabled-')]")
                split = ''
                for it3 in result_final:
                    decode = etree.tostring(it3, encoding='utf-8').decode('utf-8')
                    split = decode.split('\"')[1].split('-')[3]
                address_all_res.append(split)
                if address_pp:
                    name = etree.HTML(it_tree).xpath("//span[@class='name-given']/text()")
                    name_cj = etree.HTML(it_tree).xpath("//span[@class='aka']/text()")
                    for name_ in name:
                        if temp1 in name_:
                            res_names.append(split)
                            break
                    for name_cj_ in name_cj:
                        if temp1 in name_cj_:
                            res_names_cj.append(split)
                            break
                else:
                    address = etree.HTML(it_tree).xpath("//p[contains(@class,'address-current')]/a/text()")
                    if len(address) > 0:
                        address_ok = False
                        dpn_ok = False
                        for it2 in address:
                            address_str = it2.lower()
                            split_dpn = address_str.split(' ')[0]
                            if str(split_dpn).isdigit():
                                if int(split_dpn) == door_plate_num:
                                    dpn_ok = True
                            if data2.lower() in address_str:
                                address_ok = True

                        if address_ok:
                            address_ok_res.append(split)
                            if dpn_ok:
                                address_dpn_res.append(split)
                        else:
                            address_cj = etree.HTML(it_tree).xpath("//p[contains(@class,'address-previous')]/a/text()")
                            address_cj_ = False
                            dpn_cj_ = False
                            for it2 in address_cj:
                                address_str = it2.lower()
                                split_dpn = address_str.split(' ')[0]
                                if str(split_dpn).isdigit():
                                    if int(split_dpn) == door_plate_num:
                                        dpn_cj_ = True
                                if data2.lower() in address_str:
                                    address_cj_ = True
                                    break
                            if address_cj_:
                                address_cj_res.append(split)
                                if dpn_cj_:
                                    address_cj_dpn_res.append(split)

            result_f = ''
            temp_address_all_ = []
            for address_all_ in address_all_res:
                if len(address_all_) > 3:
                    temp_address_all_.append(address_all_)
            address_all_res = temp_address_all_
            if address_pp:
                state = -1
                if len(res_names) > 0:
                    state = 1
                    result_f = array_2_str(res_names)
                elif len(res_names_cj) > 0:
                    state = 2
                    result_f = res_names_cj[0]
                elif len(address_all_res) > 0:
                    if len(address_all_res) == 1:
                        state = -2
                    result_f = address_all_res[0]
                print()
                df_data.loc[i, "pid"] = result_f
                if state == 1:
                    df_data.loc[i, "匹配等级"] = "唯一姓名匹配"
                elif state == 2:
                    df_data.loc[i, "匹配等级"] = "曾经姓名匹配"
                elif state == -2:
                    df_data.loc[i, "匹配等级"] = "姓名-不匹配-但唯一"
                else:
                    df_data.loc[i, "匹配等级"] = "不匹配"
            else:
                # print(address_ok_res)
                # print(address_cj_res)
                # print(address_all_res)
                state = -1
                if len(address_dpn_res) > 0:
                    state = 1
                    result_f = array_2_str(address_dpn_res)
                elif len(address_cj_dpn_res) > 0:
                    state = 2
                    result_f = array_2_str(address_cj_dpn_res)
                elif len(address_ok_res) > 0:
                    state = 4
                    if len(address_ok_res) == 1:
                        state = 3
                    result_f = address_ok_res[0]
                elif len(address_cj_res) > 0:
                    state = 5
                    result_f = address_cj_res[0]
                elif len(address_all_res) > 0:
                    if len(address_all_res) == 1:
                        state = -2
                    result_f = address_all_res[0]

                print()
                df_data.loc[i, "pid"] = result_f
                if state == 1:
                    df_data.loc[i, "匹配等级"] = "唯一门牌匹配"
                elif state == 2:
                    df_data.loc[i, "匹配等级"] = "曾经门牌匹配"
                elif state == 3:
                    df_data.loc[i, "匹配等级"] = "唯一县城匹配"
                elif state == 4:
                    df_data.loc[i, "匹配等级"] = "县城匹配"
                elif state == 5:
                    df_data.loc[i, "匹配等级"] = "曾经县城匹配"
                elif state == -2:
                    df_data.loc[i, "匹配等级"] = "门牌&县城-不匹配-但唯一"
                else:
                    df_data.loc[i, "匹配等级"] = "不匹配"
            sleep(int(sleep_time))
        mkdir(procedure_file_name)
        df_data.to_excel(procedure_file_name + 'pid_result.xlsx')


def get_net_data_f(file_name_f):
    with urllib3.ProxyManager(proxy_host) as http:
        df_data = pd.read_excel(io=procedure_file_name + file_name_f)
        zd_task = False
        count = 0
        for i in range(len(df_data)):
            if zd_task:
                break
            print(str(i))
            data1 = str(df_data.loc[i, "买家姓名"]).replace(u'\xa0', ' ').split('/')[0]
            data2 = str(df_data.loc[i, "pid"]).replace(u'\xa0', ' ')
            if data2 is None or data2 == 'None' or data2 == '404' or data2 == 'nan' or data2 == '' or not isinstance(
                    data2, str):
                df_data.loc[i, "手机"] = "None"
                df_data.loc[i, "手机数据年份"] = "None"
                df_data.loc[i, "邮箱"] = "None"
                df_data.loc[i, "手机号类型"] = "None"
                continue
            if '二次地址_' not in str(df_data.loc[i, "匹配方式"]) and file_name_f == 'pid_result2.xlsx':
                print('无需二次获取detail')
                continue
            replace = data1.lower().replace('.', '').replace(' ', '-')
            data2 = data2.lower()
            result_wireless_phone = []
            result_wireless_year = []
            wireless_phone_index = []
            result_phone = []
            result_phone_year = []
            result_email = []
            pids = data2.split(',')
            need_sleep = False
            len1 = len(pids)
            count_pid = 0
            for pid in pids:
                url = url_home + '/detail/' + replace + '/' + pid
                print(url)
                if count % int(one_open_browser_count) == 0 and not need_sleep:
                    open_browser_get_cookie(url)
                headers = {
                    'User-Agent': user_agent,
                    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-',
                    'Accept-Language': 'zh-CN,zh;q=0.9',
                    'Connection': 'keep-alive',
                    'Cookie': cookie,
                }
                count = count + 1
                response = http.request('GET', url, headers=headers)
                if response.status == 403:
                    df_data.to_excel('zd_result.xlsx')
                    print('出现' + str(response.status) + '错误，任务中断')
                    zd_task = True
                    break

                if response.status != 200:
                    print('出现' + str(response.status) + '错误')
                    df_data.loc[i, "手机"] = "404"
                    df_data.loc[i, "手机数据年份"] = "404"
                    df_data.loc[i, "邮箱"] = "404"
                    df_data.loc[i, "手机号类型"] = "None"
                    continue

                tree = etree.HTML(response.data)
                # email = tree.xpath('/html/body/div[3]/div[3]/div/div[2]/div/div[2]/div[7]/div/h3/a')
                # phone = tree.xpath('/html/body/div[3]/div[3]/div/div[2]/div/div[2]/div[4]/div/h3[1]/a')
                div_all = tree.xpath("//div[@class='col-md-12 text-secondary']")
                div_phone = None
                div_email = None

                for it in div_all:
                    it_tree = etree.tostring(it, encoding='utf-8').decode('utf-8')
                    it_html = etree.HTML(it_tree)
                    phone = it_html.xpath("//span[@class='phones-label section-label']")
                    email = it_html.xpath("//i[@class='fad fa-at text-warning mr-1 mb-1']")
                    if len(phone) > 0:
                        div_phone = it_html
                    if len(email) > 0:
                        div_email = it_html

                if div_phone is not None:
                    phones = div_phone.xpath("//a[@class='phone']/text()")
                    phone_years = div_phone.xpath("//span[@class='d-block last-reported']/text()")
                    phone_types = div_phone.xpath("//span[@class='d-block phone-type']/text()")
                    for wireless_index in range(len(phone_types)):
                        phone_type = phone_types[wireless_index]
                        if "wireless" == phone_type.lower():
                            wireless_phone_index.append(wireless_index)

                    for phone1 in phones:
                        result_phone.append(phone1)
                    for phone_year in phone_years:
                        # phone_year_str = etree.tostring(phone_year, encoding='utf-8').decode('utf-8')
                        split = phone_year.split(' ')
                        result_phone_year.append(int(split[len(split) - 1]))

                if div_email is not None:
                    emails = div_email.xpath("//a[@class='email']")
                    for email1 in emails:
                        title_split = email1.attrib.get('title').split(" ")
                        result_email.append(title_split[len(title_split) - 1])

                if len(wireless_phone_index) > 0:
                    if len(wireless_phone_index) == 1:
                        result_wireless_phone.append(result_phone[wireless_phone_index[0]])
                        result_wireless_year.append(result_phone_year[wireless_phone_index[0]])
                    else:
                        temp_phone_year = []
                        for k in wireless_phone_index:
                            temp_phone_year.append(result_phone_year[k])
                        temp_index = temp_phone_year.index(max(temp_phone_year))
                        index_ = wireless_phone_index[temp_index]
                        result_wireless_phone.append(result_phone[index_])
                        result_wireless_year.append(result_phone_year[index_])
                if count_pid < len1 - 1:
                    print('合并搜索' + str(i))
                    sleep(int(sleep_time))
                need_sleep = True

            if len(result_phone_year) > 0:
                if len(result_wireless_phone):
                    df_data.loc[i, "手机"] = result_wireless_phone[result_phone_year.index(max(result_phone_year))]
                    df_data.loc[i, "手机数据年份"] = str(max(result_phone_year))
                    df_data.loc[i, "手机号类型"] = "Wireless"
                else:
                    index = result_phone_year.index(max(result_phone_year))
                    df_data.loc[i, "手机"] = result_phone[index]
                    df_data.loc[i, "手机数据年份"] = str(result_phone_year[index])
                    df_data.loc[i, "手机号类型"] = "Other"
            else:
                df_data.loc[i, "手机"] = "None"
                df_data.loc[i, "手机数据年份"] = "None"
                df_data.loc[i, "手机号类型"] = "None"

            if len(result_email) > 0:
                df_data.loc[i, "邮箱"] = array_2_str(result_email)
            else:
                df_data.loc[i, "邮箱"] = "None"
            sleep(int(sleep_time))
        if file_name_f == 'pid_result.xlsx':
            mkdir(result_file_name)
            df_data.to_excel(result_file_name + file_output_name)
        else:
            mkdir(result_file_name)
            df_data.to_excel(result_file_name + second_search_file_output_name)


def trim(str_v):
    str_v = re.sub(u"\\(.*?\\)", " ", re.sub(u"\\s\\(.*?\\)\\s", " ", str_v)).replace('  ', ' ').lower()
    if str_v.startswith(' ') or str_v.endswith(' '):
        return re.sub(r"^(\s+)|(\s+)$", "", str_v)
    return str_v


def get_url_by_line(df_data: DataFrame, line_num):
    line_num = line_num - 2
    data1 = trim(df_data.loc[line_num, "买家姓名"].replace(u'\xa0', ' '))
    data1 = re.sub(r"\d+", '', data1).split('/')[0]
    data2 = trim(df_data.loc[line_num, "县/城市"].replace(u'\xa0', ' ')).split('/')[0]
    data3 = trim(df_data.loc[line_num, "州/地区"].replace(u'\xa0', ' ')).split('/')[0]
    data4 = trim(df_data.loc[line_num, "收件人"].replace(u'\xa0', ' '))
    data4 = re.sub(r"\d+", '', data4).split('/')[0]
    data5 = trim(df_data.loc[line_num, "地址"].replace(u'\xa0', ' ')).split('/')[0]
    if data1 is None or data1 == '' or data1 == 'None' or data1 == 'nan' or not isinstance(data1, str) \
            or not isinstance(data3, str):
        print("信息无效")
        return
    temp1 = data1.lower().replace('.', '').replace(',', '')
    temp4 = data4.lower().replace('.', '').replace(',', '')
    search_name = temp1
    if data4 is not None and temp1 in temp4:
        search_name = temp4

    if len(temp1.split(' ')[0]) == 1:
        search_name = temp4

    address_pp = False
    if len(temp1.split(' ')) == 1 and search_name == temp1:
        address_pp = True

    replace = re.sub(u"\\(.*?\\)", " ", re.sub(u"\\s\\(.*?\\)\\s", " ", search_name)) \
        .replace('  ', ' ').replace(' ', '-').lower()
    # print(replace + '\t' + data2.lower().replace(' ', '-') + '\t' + data3.lower())
    if address_pp:
        url = url_home + '/address/' + data5.replace('  ', ' ').replace(' ', '-').lower() + '/' + data2.lower() \
            .replace(' ', '-') + '/' + data3.lower().replace(' ', '-')
    else:
        url = url_home + '/people/' + replace + '/' + data3.lower() + '/' + data2.lower().replace(' ', '-')
    print(url)


def name_match(name1, name2):
    split_name1 = trim(name1).split(' ')
    len1 = len(split_name1)
    count1 = 0
    for item1 in split_name1:
        if item1 in name2:
            count1 = count1 + 1
    state = 3
    if count1 > 0:
        state = 2
    if len(name2.split(' ')) > 2 and count1 == 2:
        state = 1
    if count1 == len1:
        state = 0
    return state


def array_2_str(array):
    str = ''
    for i in range(len(array)):
        if i == len(array) - 1:
            str = str + array[i]
        else:
            str = str + array[i] + ','
    return str


def array_append(array, value):
    if array is None:
        array = []
    if len(array) <= 0:
        array.append(value)
    else:
        is_append = True
        for item in array:
            if value.lower() == item.lower():
                is_append = False
        if is_append:
            array.append(value)
    return array


def address_search_2():
    print('二次搜索')
    with urllib3.ProxyManager(proxy_host) as http:
        df_data = pd.read_excel(io=result_file_name + second_search_source_file_name)
        columns_ = df_data.columns.tolist()
        columns_.insert(19, '原匹配等级')
        df_data.reindex(columns=columns_)
        count = 0
        for i in range(len(df_data)):
            print(str(i))
            ypp = df_data.loc[i, "匹配等级"]
            data1 = trim(str(df_data.loc[i, "买家姓名"]).replace(u'\xa0', ' '))
            data1 = re.sub(r"\d+", '', data1)
            data2 = trim(str(df_data.loc[i, "县/城市"]).replace(u'\xa0', ' '))
            data3 = trim(str(df_data.loc[i, "州/地区"]).replace(u'\xa0', ' '))
            data4 = trim(str(df_data.loc[i, "收件人"]).replace(u'\xa0', ' '))
            data4 = re.sub(r"\d+", '', data4)
            data5 = trim(str(df_data.loc[i, "地址"]).replace(u'\xa0', ' '))
            data1 = data1.lower().replace('.', '').replace(',', '').replace('-', '')
            data4 = data4.lower().replace('.', '').replace(',', '').replace('-', '')
            if data1 is None or data1 == '' or data1 == 'None' or data1 == 'nan' or data5 is None or not isinstance(
                    data1, str) \
                    or not isinstance(data3, str) or not isinstance(data5, str):
                continue
            match_level = df_data.loc[i, "匹配等级"]
            need_address_search = False
            if match_level == '不匹配' or match_level == '县城匹配' or match_level == '唯一县城匹配' or '曾经' in match_level:
                need_address_search = True
            if not need_address_search:
                print('无需二次搜索')
                continue
            df_data.loc[i, "匹配方式"] = '二次地址_ERROR'
            url = url_home + '/address/' + data5.replace('  ', ' ').replace(' ', '-').lower() + '/' + data2.lower() \
                .replace(' ', '-') + '/' + data3.lower().replace(' ', '-')
            print(url)
            if count % int(one_open_browser_count) == 0:
                open_browser_get_cookie(url)
            headers = {
                'User-Agent': user_agent,
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-',
                'Accept-Language': 'zh-CN,zh;q=0.9',
                'Connection': 'keep-alive',
                'Cookie': cookie,
            }
            count = count + 1
            response = http.request('GET', url, headers=headers)
            if response.status == 403:
                print(cookie)
                print(headers)
                df_data.to_excel('zd_pid_result.xlsx')
                print('出现' + str(response.status) + '错误，任务中断')
                break
            if response.status != 200:
                print('出现' + str(response.status) + '错误')
                continue
            tree = etree.HTML(response.data)
            table = tree.xpath("//div[@class='card card-hover']")
            pid_all_result = []
            name_match_0 = []
            name_match_1 = []
            name_match_2 = []
            name_match_s_0 = []
            name_match_s_1 = []
            name_match_s_2 = []
            for it in table:
                # address = it.xpath("//p[contains(@class,'address-current')]")
                it_tree = etree.tostring(it, encoding='utf-8').decode('utf-8').lower()
                result_final = etree.HTML(it_tree).xpath("//div[starts-with(@id,'opt-out-disabled-')]")
                split = ''
                for it3 in result_final:
                    decode = etree.tostring(it3, encoding='utf-8').decode('utf-8')
                    split = decode.split('\"')[1].split('-')[3]
                pid_all_result.append(split)
                name = etree.HTML(it_tree).xpath("//span[@class='name-given']/text()")
                name_cj = etree.HTML(it_tree).xpath("//span[@class='aka']/text()")
                for name_ in name:
                    match_data1 = name_match(data1, name_)
                    match_data4 = name_match(data4, name_)
                    if match_data1 == 0:
                        name_match_0 = array_append(name_match_0, split)
                    elif match_data1 == 1:
                        name_match_1 = array_append(name_match_1, split)
                    elif match_data1 == 2:
                        name_match_2 = array_append(name_match_2, split)

                    if match_data4 == 0:
                        name_match_s_0 = array_append(name_match_s_0, split)
                    elif match_data4 == 1:
                        name_match_s_1 = array_append(name_match_s_1, split)
                    elif match_data4 == 2:
                        name_match_s_2 = array_append(name_match_s_2, split)
                for name_cj_ in name_cj:
                    match_data1 = name_match(data1, name_cj_)
                    match_data4 = name_match(data4, name_cj_)
                    if match_data1 == 0:
                        name_match_0 = array_append(name_match_0, split)
                    elif match_data1 == 1:
                        name_match_1 = array_append(name_match_1, split)
                    elif match_data1 == 2:
                        name_match_2 = array_append(name_match_2, split)
                    if match_data4 == 0:
                        name_match_s_0 = array_append(name_match_s_0, split)
                    elif match_data4 == 1:
                        name_match_s_1 = array_append(name_match_s_1, split)
                    elif match_data4 == 2:
                        name_match_s_2 = array_append(name_match_s_2, split)
            result_f = ''
            if len(name_match_0) > 0:
                result_f = array_2_str(name_match_0)
                df_data.loc[i, "匹配等级"] = '姓名全匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_买家姓名'
            elif len(name_match_s_0) > 0:
                result_f = array_2_str(name_match_s_0)
                df_data.loc[i, "匹配等级"] = '姓名全匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_收件人'
            elif len(name_match_1) > 0:
                result_f = array_2_str(name_match_1)
                df_data.loc[i, "匹配等级"] = '姓名部分匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_买家姓名'
            elif len(name_match_s_1) > 0:
                result_f = array_2_str(name_match_s_1)
                df_data.loc[i, "匹配等级"] = '姓名部分匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_收件人'
            elif len(name_match_2) > 0:
                result_f = name_match_2[0]
                df_data.loc[i, "匹配等级"] = '姓名部分匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_买家姓名'
            elif len(name_match_s_2) > 0:
                result_f = name_match_s_2[0]
                df_data.loc[i, "匹配等级"] = '姓名部分匹配'
                df_data.loc[i, "匹配方式"] = '二次地址_收件人'
            if result_f != '':
                df_data.loc[i, "pid"] = result_f
            sleep(int(sleep_time))
        mkdir(procedure_file_name)
        df_data.to_excel(procedure_file_name + 'pid_result2.xlsx')


def start_work():
    get_net_data(file_name)
    if os.path.exists(procedure_file_name + 'pid_result.xlsx'):
        get_net_data_f('pid_result.xlsx')
    else:
        print("需要的前置文件未找到")


def mkdir(path):
    folder = os.path.exists(path)

    if not folder:  # 判断是否存在文件夹如果不存在则创建为文件夹
        os.makedirs(path)  # makedirs 创建文件时如果路径不存在会创建这个路径
        print


lg = logging.getLogger("Error")


def init_log():
    log_path = os.getcwd() + "/log"
    try:
        if not os.path.exists(log_path):
            os.makedirs(log_path)
    except:
        print("创建日志目录失败")
        exit(1)
    if len(lg.handlers) == 0:  # 避免重复
        # 2.创建handler(负责输出，输出到屏幕stream-handler,输出到文件filehandler)
        filename = os.path.join(log_path, 'project.log')
        fh = logging.FileHandler(filename, mode="a", encoding="utf-8")  # 默认mode 为a模式，默认编码方式为utf-8
        sh = logging.StreamHandler()
        # 3.创建formatter：
        formatter = logging.Formatter(
            fmt='%(asctime)s - %(levelname)s - Model:%(filename)s - Fun:%(funcName)s - Message:%(message)s - Line:%('
                'lineno)d')
        # 4.绑定关系：①logger绑定handler
        lg.addHandler(fh)
        lg.addHandler(sh)
        # # ②为handler绑定formatter
        fh.setFormatter(formatter)
        sh.setFormatter(formatter)
        # # 5.设置日志级别(日志级别两层关卡必须都通过，日志才能正常记录)
        lg.setLevel(40)
        fh.setLevel(40)
        sh.setLevel(40)


file_name_ = 'work.xlsx'
proxy_host_ = 'http://127.0.0.1:33210/'
sleep_time_ = '1'
file_output_name_ = 'result.xlsx'
one_open_browser_count_ = '500'
second_search_file_output_name_ = 'result.xlsx'
second_search_source_file_name_ = 'result_second_search.xlsx'


def main_method():
    while True:
        s = input('\n输入数字：\n1. 开始根据提供的“' + file_name + '“进行搜索，产生的结果将会在“' + file_output_name +
                  '”中，按 10 搜索完直接进入二次搜索\n2. 输入对应的excel行数生成对应的网址\n3. 对“' + second_search_source_file_name +
                  '"文件进行二次搜索（地址搜索），产生的结果将会在“' + second_search_file_output_name + '中\n')
        if s.isdigit():
            if int(s) == 1:
                if os.path.exists(file_name):
                    # 开始工作
                    start_work()
                else:
                    print("在该目录下没有找到名为" + str(file_name) + '的文件')
            if int(s) == 2:
                if os.path.exists(file_name):
                    excel = pd.read_excel(file_name)
                    while True:
                        excel_num = input('输入excel行数: ')
                        if excel_num.isdigit():
                            next_ = True
                            if int(excel_num) - 2 >= len(excel):
                                print('输入的行数太大了，超过有效行数')
                                next_ = False
                            if int(excel_num) - 2 < 0:
                                print('输入的行数太小了，并且不能输入标题行')
                                next_ = False
                            if next_:
                                break
                        else:
                            print('输入的必须只能是数字')
                    # 获取指定  行数  的网站链接，如果对结果产生疑问，可以用以下方法直接生成链接查看
                    get_url_by_line(excel, int(excel_num))
                else:
                    print("在该目录下没有找到名为" + str(file_name) + '的文件')
            if int(s) == 3:
                if os.path.exists(second_search_source_file_name):
                    # 开始二次搜索
                    address_search_2()
                    if os.path.exists(procedure_file_name + 'pid_result2.xlsx'):
                        get_net_data_f('pid_result2.xlsx')
                    else:
                        print("二次搜索需要的前置文件未找到")
                else:
                    print("在该目录下没有找到名为" + str(second_search_source_file_name) + '的文件')
            if int(s) == 10:
                if os.path.exists(file_name):
                    # 开始工作
                    start_work()
                    if os.path.exists(result_file_name + second_search_source_file_name):
                        # 开始二次搜索
                        address_search_2()
                        if os.path.exists(procedure_file_name + 'pid_result2.xlsx'):
                            get_net_data_f('pid_result2.xlsx')
                        else:
                            print("二次搜索需要的前置文件未找到")
                else:
                    print("在该目录下没有找到名为" + str(file_name) + '的文件')
        else:
            print('输入的必须只能是数字')

        print('按esc退出程序，按tab键继续')
        # 检测是否按下了esc键
        while True:
            if keyboard.is_pressed('esc'):
                quit()  # 退出循环
            if keyboard.is_pressed('tab'):
                break


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    init_log()
    print_hi('欢迎使用，根据提示完成所需操作')
    if file_name is None or file_name.replace(' ', '') == '':
        print('配置文件file_name为空，默认使用' + file_name_ + '， 请将文件名改为' + file_name_ + '， 或者修改配置文件')
        file_name = file_name_
    if not file_name.endswith('.xlsx') and not file_name.endswith('.xls'):
        print('配置文件file_name后缀（格式）不对，默认使用' + file_name_ + '， 请将文件名改为' + file_name_ + '， 或者修改配置文件')
        file_name = file_name_
    if proxy_host is None or proxy_host.replace(' ', '') == '':
        print('配置文件host为空，默认使用' + proxy_host_ + '， 请修改配置文件，否则将影响正常运行')
        proxy_host = proxy_host_
    if sleep_time is None or sleep_time.replace(' ', '') == '':
        print('配置文件sleep_time为空，默认使用' + sleep_time_ + '， 请修改配置文件，否则将影响正常运行')
        sleep_time = sleep_time_
    if not sleep_time.isdigit():
        print('配置文件sleep_time不是纯数字，默认使用' + sleep_time_ + '， 请修改配置文件，否则将影响正常运行')
        sleep_time = sleep_time_
    if file_output_name is None or file_output_name.replace(' ', '') == '':
        print('配置文件file_name为空，默认使用' + file_output_name_ + '， 请将文件名改为' + file_output_name_ + '， 或者修改配置文件')
        file_output_name = file_output_name_
    if not file_output_name.endswith('.xlsx') and not file_output_name.endswith('.xls'):
        print('配置文件file_name后缀（格式）不对，默认使用' + file_output_name_ + '， 请将文件名改为' + file_output_name_ + '， 或者修改配置文件')
        file_output_name = file_output_name_
    if second_search_file_output_name is None or second_search_file_output_name.replace(' ', '') == '':
        print(
            '配置文件file_name为空，默认使用' + second_search_file_output_name_ + '， 请将文件名改为' + second_search_file_output_name_ + '， 或者修改配置文件')
        second_search_file_output_name = second_search_file_output_name_
    if not second_search_file_output_name.endswith('.xlsx') and not second_search_file_output_name.endswith('.xls'):
        print(
            '配置文件file_name后缀（格式）不对，默认使用' + second_search_file_output_name_ + '， 请将文件名改为' + second_search_file_output_name_ + '， 或者修改配置文件')
        second_search_file_output_name = second_search_file_output_name_
    if second_search_source_file_name is None or second_search_source_file_name.replace(' ', '') == '':
        print(
            '配置文件file_name为空，默认使用' + second_search_source_file_name_ + '， 请将文件名改为' + second_search_source_file_name_ + '， 或者修改配置文件')
        second_search_source_file_name = second_search_source_file_name_
    if not second_search_source_file_name.endswith('.xlsx') and not second_search_source_file_name.endswith('.xls'):
        print(
            '配置文件file_name后缀（格式）不对，默认使用' + second_search_source_file_name_ + '， 请将文件名改为' + second_search_source_file_name_ + '， 或者修改配置文件')
        second_search_source_file_name = second_search_source_file_name_
    if one_open_browser_count is None or one_open_browser_count.replace(' ', '') == '':
        print('配置文件one_open_browser_count为空，默认使用' + one_open_browser_count_ + '， 请修改配置文件，否则将影响正常运行')
        one_open_browser_count = one_open_browser_count_
    if not one_open_browser_count.isdigit():
        print('配置文件one_open_browser_count不是纯数字，默认使用' + one_open_browser_count_ + '， 请修改配置文件，否则将影响正常运行')
        one_open_browser_count = one_open_browser_count_
    try:
        main_method()
        pass
    except Exception as e:
        lg.error(e)
