import requests
import hashlib
import base64
import time
import random
import pandas as pd
import html

def get_maoyan_data():
    url = 'https://piaofang.maoyan.com/dashboard-ajax'
    user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

    # 参数构造
    timestamp = str(int(time.time() * 1000))
    index = str(random.randint(100, 999))
    user_agent_encoded = base64.b64encode(user_agent.encode()).decode()

    # 生成签名
    content = f"method=GET&timeStamp={timestamp}&User-Agent={user_agent_encoded}&index={index}&channelId=40009&sVersion=2&key=A013F70DB97834C0A5492378BD76C53A"
    sign = hashlib.md5(content.encode()).hexdigest()

    # 请求头与参数
    headers = {
        'User-Agent': user_agent,
        'Referer': 'https://piaofang.maoyan.com/dashboard'
    }
    params = {
        'timeStamp': timestamp,
        'User-Agent': user_agent_encoded,
        'index': index,
        'signKey': sign,
        'channelId': '40009',
        'sVersion': '2'
    }

    response = requests.get(url, headers=headers, params=params)
    data = response.json()
    return data


def extract_nine_tail_dragon_data(data):
    movie_list = data.get('movieList', {}).get('data', {}).get('list', [])
    if not movie_list:
        return None

    for movie in movie_list:
        movie_info = movie.get('movieInfo', {})
        movie_name = movie_info.get('movieName', '')
        decoded_movie_name = html.unescape(movie_name)

        if decoded_movie_name == '哪吒之魔童闹海':
            sum_box_desc = movie.get('sumBoxDesc', '')
            timestamp = time.strftime('%Y-%m-%d %H:%M', time.localtime())
            return f"哪吒之魔童闹海 - {timestamp} - 总票房{html.unescape(sum_box_desc)}"

    return None


def main():
    while True:
        data = get_maoyan_data()
        nine_tail_dragon_data = extract_nine_tail_dragon_data(data)
        if nine_tail_dragon_data:
            print(nine_tail_dragon_data)
        else:
            print("未找到《哪吒之魔童闹海》的数据")
        time.sleep(300)  # 每五分钟抓取一次数据


if __name__ == "__main__":
    main()
