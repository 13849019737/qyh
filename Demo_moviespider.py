import re
import requests
import xlwt
import xlrd
from xlutils.copy import copy
import time


class MovieFound(object):
    """爬取电影名称和评分"""

    # 爬取次数
    __count = 0

    def __init__(self, url):

        # 传入网址
        self.response = requests.get(url)

        # 解码信息
        self.result = self.response.content.decode("utf-8")

        # 获取特点内容
        self.movieName = re.findall("<h3><.*>(.*?)</a><span>(.*?)</span></h3>", self.result)

        # 记录爬取次数
        MovieFound.__count += 1

    def add_excel(self):
        """添加到表格"""

        # 判断是否是第一次爬取，是则创建新表格，否则追加电影信息
        if MovieFound.__count == 1:

            # 创建新表格
            excel = xlwt.Workbook(encoding="utf-8")
            sheet = excel.add_sheet("电影评分排行", cell_overwrite_ok=True)

            # 设置表头
            sheet.write(0, 0, "电影名")
            sheet.write(0, 1, "电影评分")

            # 导入内容
            for i, j in enumerate(self.movieName):
                for x, y in enumerate(j):
                    sheet.write(i+1, x, y)

            # 保存表格
            excel.save("电影排行.xls")

        else:

            # 打开表格
            file = xlrd.open_workbook("电影排行.xls")
            in_sheet = file.sheets()[0]

            # 查看总行数
            rows = in_sheet.nrows

            # 建立副本
            new_file = copy(file)
            sheet = new_file.get_sheet(0)

            # 追加电影信息
            for i, j in enumerate(self.movieName):
                for x, y in enumerate(j):
                    sheet.write(i+rows, x, y)

            # 覆盖原表格
            new_file.save("电影排行.xls")

    @classmethod
    def get_count(cls):
        """查看爬取次数"""

        return cls.__count

    @staticmethod
    def sort_excel():
        """将表格中电影评分排序"""

        # 打开表格
        file = xlrd.open_workbook("电影排行.xls")
        in_sheet = file.sheets()[0]

        # 查看总行数
        rows = in_sheet.nrows

        # 读取每行电影名和评分
        row_info = in_sheet.row_values(1)
        movie_list = []

        # 把电影名和评分以元组的形式读取
        for i in range(1, rows):
            row_info = in_sheet.row_values(i)
            movie_list.append(tuple(row_info))

        # 将元组列表按照第二个元素评分，降序排列
        sort_info = sorted(movie_list, key=lambda a: a[1], reverse=True)

        # 创建副本
        new_file = copy(file)
        sheet = new_file.get_sheet(0)

        # 将按照评分排序的内容覆盖原内容
        for i, j in enumerate(sort_info):
            for x, y in enumerate(j):
                sheet.write(i+1, x, y)

        # 覆盖原表格
        new_file.save("电影排行.xls")


if __name__ == '__main__':

    # 爬取电影页中所有页数电影信息
    for page in range(1, 329):

        # 不同页数的url
        url = f"https://www.pianku.tv/mv/------{page}.html"

        # 创建对象
        movie = MovieFound(url)

        # 将电影信息保存到表格
        movie.add_excel()

        # 排序表格
        movie.sort_excel()

        # 查看爬取次数
        print(movie.get_count())

        # 设置刷新时间
        time.sleep(2)
