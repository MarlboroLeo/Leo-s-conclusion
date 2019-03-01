import json
import random
import re
import threading
import time
from queue import Queue

import requests
from openpyxl import load_workbook


class JdCommentsSpider(object):
    """京东商品评论爬取"""
    def __init__(self, cookie_str, product_id, call_back, sort_type="5", ):
        self.url_queue = Queue()  # url队列
        self.comments_queue = Queue()  # 评论数据队列
        self.headers = {
            "referer": "https://item.jd.com/5089253.html",
            "user - agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko)Chrome/70.0.3538.67 Safari/537.36",
            'Connection': 'close',
            'cookie': cookie_str
        }
        self.product_id = product_id
        self.sort_type = sort_type
        self.call_back = call_back
        self.wb = load_workbook('jd.xlsx')  # 同级目录下要存在此文件
        self.ws = self.wb["Sheet1"]

    def ua_and_proxy(self):
        """TODO User-Agent和代理IP"""
        pass

    def gen_url(self):
        first_url = "https://sclub.jd.com/comment/productPageComments.action?callback={}&productId={}&score=0&sortType={}&page={}&pageSize=10&isShadowSku=0&fold=1"
        # productId=5089253 商品ID 抓包获取
        # sortType=6为按时间排序，5为推荐排序
        for page_num in range(91):  # page共91页
            url = first_url.format(self.call_back, self.product_id, self.sort_type, page_num)
            self.url_queue.put(url)

    def extra_comments(self):
        """发起请求"""
        print("{}线程request任务开始".format(threading.current_thread()))
        while self.url_queue.qsize():
            url = self.url_queue.get()
            print("正在爬取第{}页".format(re.search(r'page=(\d+)&', url).group(1)))
            time.sleep(random.randint(4, 6))
            response = requests.get(url, headers=self.headers)
            if response.status_code == 200:
                try:
                    # 提取json格式数据
                    response_json = re.match(r"fetchJSON_comment.*?\((.*)\);", response.text).group(1)
                    self.comments_queue.put(response_json)
                    self.url_queue.task_done()
                except Exception as e:
                    print(e)
                    print(response.text)
            else:
                print(response.status_code)
                pass

    def save_to_excel(self):
        """提取字段并保存至excel"""
        print("{}线程save任务开始".format(threading.current_thread()))
        while True:
            comment_json = self.comments_queue.get()
            comment_dict = json.loads(comment_json)  # 转为dict
            for comment in comment_dict.get("comments"):
                comment_id = comment.get("id")  # 评论id
                creation_time = comment.get("creationTime")  # 发布时间
                content = comment.get("content")  # 评论内容
                row = ["{}".format(comment_id), "{}".format(creation_time), content]  # 组装成列表
                self.ws.append(row)  # 写入excel表格
                self.wb.save("jd.xlsx")
            self.comments_queue.task_done()

    def run(self):
        """主逻辑"""
        thread_list = []
        self.gen_url()
        for i in range(1):
            t_extra_comments = threading.Thread(target=self.extra_comments, args=(threading .Thread.name, ), )
            thread_list.append(t_extra_comments)
            t_save = threading.Thread(target=self.save_to_excel, args=(threading .Thread.name, ))
            thread_list.append(t_save)
        for t in thread_list:
            t.setDaemon(True)  # 守护主进程
            t.start()
        for q in [self.url_queue, self.comments_queue]:
            q.join()  # 主进程阻塞
            print("主线程结束")


if __name__ == '__main__':
    print("爬取狗东商品评论 v1.0n")
    cookie = input("粘贴cookie:")
    product_id = input("product_id：")
    call_back = input("callback:")
    sort_type = input("sortType=6为按时间排序，默认5为推荐排序:")
    jd = JdCommentsSpider(cookie, product_id, call_back, sort_type)
    jd.run()
