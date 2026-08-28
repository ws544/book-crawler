import csv
from enum import Enum
import json
import re
import requests

from lxml import etree
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
from peewee import MySQLDatabase, Model
from peewee import AutoField, CharField, IntegerField, DecimalField


# 表示写入文件的类型的枚举类
class WriteToType(Enum):
    CSV = 1
    JSON = 2
    EXCEL = 3
    MYSQL = 4


headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
}
url = "https://movie.douban.com/top250"
db = MySQLDatabase(
    database='spider_db',
    host='localhost',
    port=3306,
    user='root',
    password='123456')


# 电影基础信息模型（表：top250_douban）
class Movie(Model):
    id = AutoField(primary_key=True)
    rank = IntegerField(unique=True)
    title = CharField(max_length=100, unique=True)
    score = DecimalField(max_digits=2, decimal_places=1)
    year = IntegerField()
    rating_count = IntegerField()

    class Meta:
        database = db
        db_table = "top250_douban"


# 电影导演信息模型（表：top250_douban_directors）
class MovieDirectors(Model):
    id = AutoField(primary_key=True)
    movie_id = IntegerField()
    director = CharField(max_length=100)

    class Meta:
        database = db
        db_table = "top250_douban_directors"


# 储存所有电影信息到CSV文件
def write_movies_to_csv(movies):
    with open('douban_top250_movies.csv', 'w', newline='', encoding='utf-8-sig') as f:
        fieldnames = ["排名", "标题", "评分", "年份", "评价人数", "导演"]
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(movies)
    print("电影信息已写入 douban_top250_movies.csv")


# 储存所有电影信息到JSON文件
def write_movies_to_json(movies):
    with open('douban_top250_movies.json', 'w', encoding='utf-8') as f:
        json.dump(movies, f, ensure_ascii=False, indent=2)
    print("电影信息已写入 douban_top250_movies.json")


# 储存所有电影信息到Excel文件
def write_movies_to_excel(movies):
    wb = Workbook()
    ws = wb.active
    ws.title = "豆瓣电影Top250"
    excel_headers = ["排名", "标题", "评分", "年份", "评价人数", "导演"]
    thin_side = Side(style='thin')
    for col_num, header in enumerate(excel_headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)
        cell.border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    for movie in movies:
        ws.append(list(movie.values()))

    wb.save('douban_top250_movies.xlsx')
    print("电影信息已写入 douban_top250_movies.xlsx")


# 储存所有电影信息到MySQL
def write_movies_to_mysql(movies):
    try:
        db.connect()
    except Exception as e:
        print(f"数据库连接失败: {e}")
        return

    Movie.delete().execute()
    MovieDirectors.delete().execute()
    db.execute_sql("ALTER TABLE top250_douban AUTO_INCREMENT=1;")
    db.execute_sql("ALTER TABLE top250_douban_directors AUTO_INCREMENT=1;")
    print("数据删除成功")

    # 开启事务
    with db.atomic():
        for movie_dict in movies:
            movie = Movie.create(rank=movie_dict["排名"],
                                 title=movie_dict["标题"],
                                 score=movie_dict["评分"],
                                 year=movie_dict["年份"],
                                 rating_count=movie_dict["评价人数"])
            print(f"电影《{movie.title}》写入成功，id={movie.id}")

            director_list = movie_dict["导演"]
            for director_name in director_list:
                director_obj = MovieDirectors.create(movie_id=movie.id, director=director_name)
                print(f"导演{director_obj.director}写入成功，id={director_obj.id}")


# 获取电影标题
def get_title(item):
    title_elements = item.xpath('.//div[@class="hd"]/a/span[1]/text()')
    title = title_elements[0] if title_elements else None
    return title


# 获取电影评分
def get_score(item):
    score_elements = item.xpath('.//span[@class="rating_num"]/text()')
    score = score_elements[0] if score_elements else None
    return score


# 获取电影排名
def get_rank(item):
    rank_elements = item.xpath('.//div[@class="pic"]/em')
    rank_html = etree.tostring(rank_elements[0], encoding="unicode").strip() if rank_elements else None
    rank_match = re.search(r'\d+', rank_html)
    rank = rank_match.group() if rank_match else None
    return rank


# 获取电影评价人数
def get_rating_count(item):
    rating_count_elements = item.xpath('.//div[@class="bd"]//span[4]/text()')
    rating_count_text = rating_count_elements[0] if rating_count_elements else None
    rating_count_match = re.search(r'\d+', rating_count_text)
    rating_count = rating_count_match.group() if rating_count_match else None
    return rating_count


# 获取包含电影导演和年份的文本
def get_director_and_year_group_text(item):
    director_elements = item.xpath('.//div[@class="bd"]/p[1]/text()')
    director_text = director_elements[0].strip() if director_elements else None
    year_text = director_elements[1].strip() if len(director_elements) == 2 else None
    return director_text, year_text


# 获取电影年份
def get_year(year_text):
    year_match = re.search(r'\d{4}', year_text)
    year = year_match.group() if year_match else "获取年份失败"
    return year


# 获取电影导演列表
def get_director_list(director_text):
    director_match = re.search(r'导演[：:]\s*(.*?)((\s{2,}主)|(\.{3}))', director_text)
    director_text = director_match.group(1).strip() if director_match else None
    director_list = director_text.split(" / ")
    for director in director_list:
        print(f"导演: {director}")
    return director_list


# 遍历电影项列表并提取信息
def process_movie_items(movie_item_list, movies, write_to):
    for item in movie_item_list:
        # 获取标题
        title = get_title(item)
        # 获取评分
        score = get_score(item)
        # 获取排名
        rank = get_rank(item)
        # 获取评价人数
        rating_count = get_rating_count(item)
        print(f"排名={rank}, 标题={title}, 评分={score}, 评价人数={rating_count}")
        # 获取年份和导演列表
        director_and_year_group_text = get_director_and_year_group_text(item)
        year = get_year(director_and_year_group_text[1])
        director_list = get_director_list(director_and_year_group_text[0])
        # 创建电影信息字典
        movie_dict = create_movie_dict(rank, title, score, year, rating_count, director_list, write_to)
        movies.append(movie_dict)
        print(f"{movie_dict=}")
    return movies


# 获取HTML内容
def get_html_content(url_, start_=0):
    page_url = f"{url_}?start={start_}"
    response = requests.get(page_url, headers=headers)
    response.raise_for_status()
    html_content = response.text
    return html_content


# 创建电影信息字典
def create_movie_dict(rank, title, score, year, rating_count, director_list, write_to):
    if write_to == WriteToType.JSON or write_to == WriteToType.MYSQL:
        director_value = director_list
    else:
        director_value = ",".join(director_list)
    movie_dict = {
        "排名": rank,
        "标题": title,
        "评分": score,
        "年份": year,
        "评价人数": rating_count,
        "导演": director_value,
    }
    return movie_dict


def scrape_top250_movies(write_to):
    movies = []
    try:
        for start in range(0, 250, 25):
            html = get_html_content(url, start)
            tree = etree.HTML(html)
            item_list = tree.xpath('//div[@class="item"]')
            movies = process_movie_items(item_list, movies, write_to)
        if write_to == WriteToType.CSV:
            write_movies_to_csv(movies)
        elif write_to == WriteToType.JSON:
            write_movies_to_json(movies)
        elif write_to == WriteToType.EXCEL:
            write_movies_to_excel(movies)
        elif write_to == WriteToType.MYSQL:
            write_movies_to_mysql(movies)
        else:
            print("未实现的写入文件类型：", write_to)
    except requests.RequestException as e:
        print(f"请求失败: {e}")
    except Exception as e:
        print(f"发生错误: {e}")


if __name__ == '__main__':
    scrape_top250_movies(WriteToType.MYSQL)

