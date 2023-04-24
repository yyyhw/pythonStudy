import re

import requests
import openpyxl
import concurrent.futures
import json
import pandas as pd
import jieba
from collections import Counter
from openpyxl.styles import NumberFormatDescriptor
from openpyxl.styles import numbers
import xlsxwriter


def get_video_info(video_id):
    """
    获取B站视频信息
    """
    url = f"https://api.bilibili.com/x/web-interface/view?bvid={video_id}"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    json_data = response.json()
    video_title = json_data['data']['title']
    comment_count = get_comment_count(video_id)
    return video_title, comment_count


def get_comment_count(video_id):
    url = f"https://api.bilibili.com/x/v2/reply?jsonp=jsonp&pn=1&type=1&oid={video_id}"
    response = requests.get(url)
    json_data = json.loads(response.text)
    return json_data['data']['page']['count']


def get_comments(video_id, page_num):
    """
    分页获取B站视频评论
    """
    url = f"https://api.bilibili.com/x/v2/reply?&pn={page_num}&type=1&oid={video_id}&sort=0&_=1621981683583"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36"
    }
    response = requests.get(url, headers=headers)
    json_data = response.json()
    comment_list = json_data['data']['replies']
    return comment_list


def get_all_comments(video_id):
    """
    获取B站视频的所有评论
    """
    video_title, comment_count = get_video_info(video_id)
    video_title = '评论'
    print(f"视频标题：{video_title}")
    print(f"评论数：{comment_count}")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["用户名", "评论内容", "点赞数", "回复数"])
    all_comments = []
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(get_comments, video_id, i) for i in range(1, comment_count // 1000 + 2)]
        # futures = [executor.submit(get_comments, video_id, i) for i in range(1, 100 // 100 + 2)]
        for future in concurrent.futures.as_completed(futures):
            comment_list = future.result()
            for comment in comment_list:
                uname = comment['member']['uname']
                message = comment['content']['message']
                like = comment['like']
                rcount = comment['rcount']
                ws.append([uname, message, like, rcount])
                all_comments.extend(comment_list)
    wb.save(f"{video_title}.xlsx")
    print("评论数据保存完成！")
    return all_comments, video_title


# 我还想在excel中的新开一个sheet来统计所有评论中的高


def get_top_words(comments, top_n):
    words = []
    stop_words = ['的', '了', '啊', '呢', '吧', '哦', '嗯', '呀','']
    for comment in comments:
        sentence = json.dumps(comment, ensure_ascii=False)
        sentence = str(sentence)
        sentence = re.sub(r'[^\u4e00-\u9fa5]+', '', sentence)
        sentence = re.sub(r'[^\w\s]', '', sentence)  # 去除标点符号
        sentence = jieba.lcut(sentence.encode('utf-8'))  # 分词
        sentence = [word for word in sentence if word not in stop_words]  # 去除停用词
        words.extend(sentence)
    word_count = Counter(words)
    top_words = word_count.most_common(top_n)
    return top_words
# 这个是我的分词的方法，把去除无用词和去除标点符号的功能实现

def save_to_excel(top_words, filename):
    # 读取已有的 Excel 文件，如果不存在就创建一个新的
    try:
        with pd.ExcelFile(filename) as xls:
            sheet_names = xls.sheet_names
    except FileNotFoundError:
        sheet_names = []
    writer = pd.ExcelWriter(filename, engine='openpyxl')

    # 如果之前已经保存了评论信息的工作表，就把它读取出来
    if 'comments' in sheet_names:
        df_comments = pd.read_excel(filename, sheet_name='comments')
        df_comments.to_excel(writer, sheet_name='comments', index=False)

    # 将新的高频词汇数据写入新的工作表
    word_df = pd.DataFrame(top_words, columns=['word', 'count'])
    word_df.to_excel(writer, 'top_words', index=False)
    workbook = writer.book
    worksheet = writer.sheets['top_words']
    format1 = NumberFormatDescriptor('#,##0')
    worksheet.column_dimensions['B'].width = '10.0'
    worksheet.column_dimensions['B'].number_format = str(format1)
    writer.close()

    # 将新的高频词汇数据写入新的工作表
    # word_df = pd.DataFrame(top_words, columns=['word', 'count'])
    # word_df.to_excel(writer, 'top_words', index=False)
    # workbook = writer.book
    # worksheet = writer.sheets['top_words']
    # format1 = workbook.add_format({'num_format': '#,##0'})
    # worksheet.set_column('B:B', None, format1)
    # writer.save()
    #
    # # 将新的高频词汇数据写入新的工作表
    # word_df = pd.DataFrame(top_words, columns=['word', 'count'])
    # word_df.to_excel(writer, 'top_words', index=False)
    # worksheet = writer.sheets['top_words']
    # workbook = writer.book
    # format1 = workbook.add_format({'num_format': '#,##0'})
    # worksheet.set_column('B:B', None, format1)
    # writer.save()


if __name__ == '__main__':
    video_id = "BV1Sk4y1471G"
    comments, video_title = get_all_comments(video_id)
    # 获取前20个高频词汇
    top_words = get_top_words(comments, 20)
    # 将高频词汇保存到Excel中
    save_to_excel(top_words, "高频词汇.xlsx")

# D:\DDesktop\pl.xlsx
# https://www.bilibili.com/video/BV1Sk4y1471G/?vd_source=35cf14edb1c651ccf5647421265145a6
# video_url = 'https://www.bilibili.com/video/BV1Sk4y1471G'
# video_url = 'https://www.bilibili.com/video/BV1Sk4y1471G/?vd_source=35cf14edb1c651ccf5647421265145a6'
# comments = get_comments(video_url)
# for comment in comments:
#     print(comment)
