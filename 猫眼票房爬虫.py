import requests
from lxml import etree
import time
import pandas as pd


headers = {
    "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"}

def request_maoyan(year, save_path, sleep_time=1):
    '''
    year：想获取的电影年份
    demo: https://piaofang.maoyan.com/rankings/year?year=2024&limit=600&tab=1
    '''
    if not save_path.endswith('xlsx'): raise TypeError('Only support excel: xlsx file')
    time.sleep(sleep_time)
    tab = abs(int(year) - 2024) + 1
    url = f'https://piaofang.maoyan.com/rankings/year?year={year}&limit=600&tab={tab}'
    print(url)
    response = requests.get(url=url, headers=headers, timeout=10)
    if response.status_code != 200:
        raise ConnectionError(f"status code: {response.status_code} is not 200, break.")
    html = etree.HTML(response.text)
    
    
    title_xpath = "//ul[@class='row']/li[2]/p[@class='first-line']/text()"
    ids_xpath = "//ul[@class='row']/@data-com"
    show_time_xpath = "//ul[@class='row']/li[2]/p[@class='second-line']/text()"
    total_piaofan_xpath = "//ul[@class='row']/li[3]/text()"
    avg_piaofang_xpath = "//ul[@class='row']/li[4]/text()"
    avg_people_xpath = "//ul[@class='row']/li[5]/text()"

    titles = html.xpath(title_xpath)

    ids = html.xpath(ids_xpath)
    if ids: ids = list(map(lambda x: x.replace(r"hrefTo,href:'/movie/", '').replace("'", ''), ids))

    show_time = html.xpath(show_time_xpath)
    if show_time: show_time = list(map(lambda x: x.replace(r" 上映", ''), show_time))

    total_piaofang = html.xpath(total_piaofan_xpath)
    avg_piaofang = html.xpath(avg_piaofang_xpath)
    avg_people = html.xpath(avg_people_xpath)

    if titles and ids and show_time and total_piaofang and avg_people and avg_people:
        titles = ['电影'] + titles
        ids = ['id'] + ids
        show_time = ['上映时间'] + show_time
    else:
        raise f"content is null, please check network or your target url is right."

    output = [
        titles, ids, show_time, total_piaofang, avg_piaofang, avg_people
    ]
    output = pd.DataFrame(zip(*output))
    output.columns = output.iloc[0]
    output = output.iloc[1:]
    output.to_excel(save_path, index=False)

if __name__ == "__main__":
    request_maoyan('2024', '2024.xlsx')
    request_maoyan('2023', '2023.xlsx')
    request_maoyan('2022', '2022.xlsx')
    request_maoyan('2021', '2021.xlsx')