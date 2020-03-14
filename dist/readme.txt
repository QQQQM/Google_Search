# 运行方法  #
配置好config文件后直接点击 Google_Search.exe 即可

# 环境配置  #
要求安装好Google Chrome 浏览器 且 chromedriver 版本与浏览器版本对应

# config说明 # 
1、save_folder 保存文件的位置
2、driver_path 为chromedriver文件存放的位置
3、scolar_url 进行Google学术搜索的网站
4、search_url 进行作者+关键词搜索的网站（默认为必应搜索）
5、key_word_1 搜索的关键词
6、sleep_time 网页中每次操作后等待的时间（2s）
7、page_num 对关键词提取的词条数目（不是页数！）
8、year 检索时年份限制（0为不限制）
9、no_window 是否展示浏览器窗口（True为不展示，False为展示）
10、xpath 定位词条位置辅助
11、several_name指定“多”的那一部分的查找名
12、one_name指定“单”的那一部分的查找名


config文件中several_name=source，one_name=intext（可以用intex和suorce和author）
则会对several.txt中的每一项A、B、C、D 和one.txt中的一项X
进行
source:" A" intext:"X"     
source:" B" intext:"X"  
source:" C" intext:"X"
……
注意：自定义several.txt和one.txt时，最后一行不要是空行


#  2020.3.14 #

# @Author: qimeng
# @File  : readme.txt
