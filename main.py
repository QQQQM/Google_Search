#!/usr/bin/env python
# -*- coding: utf-8 -*-
# @Time : 2020/3/14 11:37
# @Author : Qi Meng
# @File : main.py
import os
import re
import time
import winsound
import xlsxwriter
from selenium import webdriver
from configparser import ConfigParser


class Google_Search:
    def __init__(self):
        config_parser = ConfigParser()
        config_parser.read('config.cfg')
        config = config_parser["default"]
        self.save_folder = config['save_folder']
        self.key_word_1 = config['key_word_1']
        self.sleep_time = config['sleep_time']
        self.page_num = config['page_num']
        self.driver_path = config['driver_path']
        self.scolar_url = config['scolar_url']
        self.search_url = config['search_url']
        self.several_name = config['several_name']
        self.one_name = config['one_name']
        self.year = config['year']
        self.xpath = config['xpath']
        self.stop_num = int(config['stop_num'])
        self.author = ['Oliphant, Travis E', 'Lutz, Mark']
        if config['no_window'] == "True":
            self.window = True
        else:
            self.window = False

    def mkdir(self):
        path_name = self.save_folder
        mkdir_reserve(path_name)

    def produce_biblist(self, author, title):
        title.clear()
        author.clear()
        bibtex = []
        sleep_time = int(self.sleep_time)
        num_max = int(self.page_num) * 1

        opt = webdriver.ChromeOptions()  # 选择为chrome浏览器
        opt.headless = self.window  # 选择为展现窗口模式
        driver = webdriver.Chrome(executable_path=self.driver_path, options=opt)  # 创建浏览器对象ss
        driver.maximize_window()  # 最大化窗口
        print("\n已成功创建浏览器对象！")
        if self.year == '0':
            driver.get(
                self.scolar_url + "/scholar?hl=zh-CN&q=" + self.key_word_1)  # 打开url对应的网页
        else:
            driver.get(
                self.scolar_url + "/scholar?hl=zh-CN&as_ylo=" + self.year + "&q=" + self.key_word_1)  # 打开url对应的网页
        time.sleep(sleep_time)  # 等待2s加载
        print("已成功打开链接！")

        home_handle = driver.current_window_handle  # 获取当下的 handle 作为基础页面home_handle
        for num in range(num_max):
            handles = driver.window_handles  # handles 获取当下所有窗口标签
            for handle in handles:  # 切换为新的窗口，以便读取信息之后关闭窗口
                if handle != home_handle:
                    driver.switch_to.window(handle)
                    home_handle = driver.current_window_handle  # home handle 记录当下句柄为家句柄
            if num != 0 and num % 10 == 0:  # google学术每页有10项，自动换页
                driver.find_element_by_xpath("//*[@id='gs_n']/center/table/tbody/tr/td[12]/a/b").click()
                # driver.find_element_by_xpath("//*[@id='gs_nm']/button[2]/span/span[1]").click()
                time.sleep(sleep_time)
                print("\n" + "已成功切换翻页！")
                home_handle = driver.current_window_handle  # home handle 记录当下句柄为家句柄

            path = "//div[@id='gs_res_ccl_mid']/div[@class='gs_r gs_or gs_scl'][@data-rp = '" + str(num) + \
                   "']//div[@class='gs_fl']/a[@class='gs_or_cit gs_nph']"  # 定位第num条的“引用”按钮的xpath
            driver.find_element_by_xpath(path).click()  # 点击“引用”按钮，弹出引用弹窗
            time.sleep(sleep_time)  # 等待2s加载 bibtex 按钮
            print("已成功打开引言弹窗！")

            driver.find_element_by_xpath("//*[@id='gs_citi']/a[1]").click()  # 点击 bibtex 按钮，弹出的新的窗口
            time.sleep(sleep_time)  # 等待2s加载 bibtex 窗口
            print("已成功打开bibtex标签！")

            handles = driver.window_handles  # handles 获取当下所有窗口标签
            for handle in handles:  # 切换为新的窗口，以便读取信息之后关闭窗口
                if handle != home_handle:
                    driver.switch_to.window(handle)
            passage = driver.find_element_by_xpath("/html/body/pre")  # 获取bibtex的具体内容
            bibtex.append(passage.text)  # 加入 bibtex列表中
            driver.back()
            # driver.close()  # 关闭 bibtex 窗口
            # driver.switch_to.window(home_handle)
            num += 1
            driver.find_element_by_xpath("//*[@id='gs_cit-x']/span[1]").click()  # 点击 取消 按钮
            time.sleep(sleep_time)  # 等待2s加载
            print("第" + str(num) + "篇bibtex提取完成！\n")
            if num % self.stop_num == 0:
                winsound.Beep(600, 1000)
                break_out = input("\n是否继续？（1继续，0退出）\n")
                if int(break_out) == 0:
                    break
        print("\n" + "bibtex 全部提取完成！")
        deal_bib(bibtex, title, author)  # 提取 bibtex 中的作者以及标题，保存为list
        driver.close()

    def bibtex(self):
        author = []
        author_all = []
        title = []
        key_word = re.sub(r"[^a-zA-Z0-9]", "_", self.key_word_1)
        key_file = self.save_folder + '/' + key_word     # key_file 存放关键词+作者的文件夹
        mkdir_reserve(key_file)                          # 创建关键词文件夹
        if len(key_word) >= 22:
            key_word = key_word[:20]
        xlsx_name = key_file + '/' + key_word + '.xlsx'  # 以关键词名字命名 xlsx 表格
        workbook = xlsxwriter.Workbook(xlsx_name)        # 建立 xlsx 表格

        work_sheet = []
        work_sheet.append(workbook.add_worksheet('总表'))
        work_sheet[0].set_column(0, 4, 50)                # 设置宽度
        title_data = ["作者", "文章标题"]                  # 设置标题文字
        work_sheet[0].write_row('A1', title_data)         # 写入title_data

        self.produce_biblist(author, title)               # 函数 1 ：调用函数产生 biblist

        author = func(author)  # 提取 author 即作者信息
        title = func(title)  # 提取 title 即标题信息

        for auth in author:  # 提取 author中的单个作者信息
            author_all.append(auth.split(" and "))
        author_all = func(author_all)  # 提取 author_all 即所有作者信息
        self.author = func(author_all)

        for item in range(len(author_all)):
            author_all[item] = author_all[item].strip()

        write_several(author_all)

        for i in range(len(author)):
            work_sheet[0].write_column('A2', author)  # 写入第一列作者信息
            work_sheet[0].write_column('B2', title)  # 写入第二列标题信息
        workbook.close()

    def author_url(self, flag):
        several_list = read_several()
        one_list = read_one()

        key_word = re.sub(r"[^a-zA-Z0-9]", "_", one_list)
        key_file = self.save_folder + '/' + key_word  # key_file 存放关键词+作者的文件夹
        author_file = key_file + "/Author"  # key_file 存放关键词+作者的文件夹
        mkdir_reserve(key_file)
        mkdir_reserve(author_file)

        if len(key_word) >= 22:
            key_word = key_word[:20]
        xlsx_name = key_file + "/Author-" + key_word + '.xlsx'  # 以关键词名字命名 xlsx 表格
        workbook_0 = xlsxwriter.Workbook(xlsx_name)  # 建立 xlsx 表格
        work_sheet_0 = workbook_0.add_worksheet("总表")
        work_sheet_0.set_column(0, 0, 20)  # 设置宽度
        work_sheet_0.set_column(1, 1, 200)  # 设置宽度
        title_data = [self.several_name, "链接"]  # 设置标题文字
        work_sheet_0.write_row('A1', title_data)  # 写入title_data

        sleep_time = int(self.sleep_time)
        opt = webdriver.ChromeOptions()                                           # 选择为chrome浏览器
        opt.headless = self.window                                                # 选择为展现窗口模式
        driver = webdriver.Chrome(executable_path=self.driver_path, options=opt)  # 创建浏览器对象ss
        driver.maximize_window()                                                  # 最大化窗口
        print("\n已成功创建浏览器对象！")

        cnt = 1
        cnt_num = 1
        ulr_list_box = []
        for author in several_list:
            # 表格初始化
            if len(author) >= 22:
                author_tem = author[:20]
            else:
                author_tem = author
            xlsx_name = author_file + "/" + author_tem + '.xlsx'  # 以关键词名字命名 xlsx 表格
            workbook = xlsxwriter.Workbook(xlsx_name)  # 建立 xlsx 表格
            work_sheet = workbook.add_worksheet(author_tem)

            work_sheet.set_column(0, 0, 20)                    # 设置宽度
            work_sheet.set_column(1, 1, 200)                   # 设置宽度
            title_data = [self.several_name, "链接"]           # 设置标题文字
            work_sheet.write_row('A1', title_data)             # 写入title_data
            work_sheet.write(1, 0, author)                     # 写入title_data
            work_sheet_0.write(cnt, 0, author)                 # 写入第二列标题信息

            search_item = self.several_name + ":\"" + author + "\" " + self.one_name + ":\"" + one_list + "\""
            print("\n检索内容：", search_item)

            if flag == 0:
                driver.get(self.search_url + "/search?q=" + search_item)       # 打开url对应的网页
            else:
                # driver.get(self.scolar_url + "/scholar?&q=" + search_item)    # 打开url对应的网页
                if cnt == 1:
                    driver.get(self.scolar_url +  "/scholar?&q= ")             # 打开url对应的网页
                else:
                    driver.execute_script("window.open('" + self.scolar_url + "/scholar?&q= " + "')")
                    handles = driver.window_handles
                    driver.switch_to.window(handles[-1])
                time.sleep(sleep_time)                                          # search_item
                driver.find_element_by_id('gs_hdr_tsi').clear()                 # 清空输入窗口
                time.sleep(sleep_time)
                driver.find_element_by_id('gs_hdr_tsi').send_keys(search_item)  # 输入窗口
                driver.find_element_by_id('gs_hdr_tsb').click()  # 点击确定
                # pyautogui.press('enter')
            time.sleep(sleep_time)  # 等待加载
            print("已成功打开链接！")
            url_box = []
            try:
                url_box.append(driver.current_url)
                ulr_list_box.append(driver.current_url)
            except:
                pass
            if flag == 1:
                self.xpath = "//h3/a"
            for link in driver.find_elements_by_xpath(self.xpath):
                print(link.get_attribute('href'))
                url_box.append(link.get_attribute('href'))
            # print(url_box)
            if len(url_box) == 0:
                url_box.append("NA")

            work_sheet.write_column('B2', url_box)  # 写入第二列标题信息
            work_sheet_0.write_column(cnt, 1, url_box)  # 写入第二列标题信息
            cnt += len(url_box)
            workbook.close()
            if cnt_num % self.stop_num == 0:
                winsound.Beep(600,1000)
                break_out = input("\n是否继续？（1继续，0退出）\n")
                if int(break_out) == 0:
                    break
            cnt_num += 1
        workbook_0.close()
        # driver.close()
        f2 = open(key_file + '/' + "url-list.bat", "w+")
        f2.write("@echo off\n")
        for item in range(len(ulr_list_box)):
            f2.write("start chrome.exe ")
            f2.write(ulr_list_box[item])
            f2.write("\n")
            f2.write("TIMEOUT /T ")
            f2.write(str(sleep_time))
            f2.write("\n")
        f2.write("pause\nexit")
        f2.close()
        print(ignoreit)


# 去除嵌套list，并且去除空的元素
def func(x):   # 去除嵌套list，并且去除空的元素
    empty_str = ['']
    a = [" " if x in empty_str else x for x in ([a for b in x for a in func(b)] if isinstance(x, list) else [x])]
    for item in range(len(a)):
        a[item] = a[item].strip()
    return a


# 此函数为对获取的 bibtex 的列表进行处理，形成可以写入excel的格式
def deal_bib(bib_list, title, author):
    for bib in bib_list:
        pattern_title = re.compile("title={(.*?)}", re.I)
        title.append(pattern_title.findall(bib)[0])
        pattern_author = re.compile("author={(.*?)}", re.I)
        author.append(pattern_author.findall(bib)[0])


# 创建文件夹
def mkdir_reserve(path_name):
    path_name = path_name.strip()  # 去除首位空格
    path_name = path_name.rstrip("/")  # 去除尾部 \ 符号
    isExists = os.path.exists(path_name)  # 判断路径是否存在.存在--True,不存在--False
    if not isExists:
        os.makedirs(path_name)
        print(path_name + ' 创建成功')
        # t.insert('end', "\n" + path_name + ' 创建成功')
        # window.update()
        return True
    else:
        print(path_name + ' 目录已存在')
        # t.insert('end', "\n" + path_name + ' 目录已存在')
        # window.update()
        return False


def write_several(a):
    f1 = open("several.txt", "w+")
    # ccc = func(f1.read().split("\n"))
    # ccc = [i for i in ccc if (len(str(i)) != 0)]
    for i in range(len(a)-1):
        f1.write(a[i])
        f1.write("\n")
    f1.write(a[-1])
    f1.close()


def read_several():
    f1 = open("several.txt", "r",encoding='utf-8')
    ccc = func(f1.read().split("\n"))
    ccc = [i for i in ccc if (len(str(i)) != 0)]
    # ccc = list(set(func(ccc)))
    ccc = sorted(list(set(ccc)),key = ccc.index)
    for item in ccc[::-1]:
        if item == "others":
            ccc.remove(item)
    return ccc


def read_one():
    f1 = open("one.txt", "r",encoding='utf-8')
    return f1.readline()


if __name__ == '__main__':
    google_search = Google_Search()
    print("指定存储位置为：",google_search.save_folder)
    google_search.mkdir()

    while(True):
        print("# 欢迎使用 Google 抓取url工具！ #")
        print("** 1：根据关键词获得作者列表。")
        print("** 2：作者列表与关键词获得url（必应）。")
        print("** 3：作者列表与关键词获得url（谷粉）。")
        print("** 0：退出。")
        option = input("# 请输入选项： #\n")
        if option == '0':
            break
        elif option == '1':
            google_search.bibtex()
        elif option == '2':
            google_search.author_url(0)
        elif option == '3':
            google_search.author_url(1)
        else:
            print("输入不合规范！")
        option = input("\n# 是否继续？（1继续，0退出） #\n")
        if option == '0':
            break
        else:
            os.system('cls')


