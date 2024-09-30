# coding=utf-8
import requests
from lxml import etree
import xlsxwriter
from concurrent.futures import ThreadPoolExecutor
from fake_useragent import UserAgent
ua = UserAgent()
import sqlite3
proxies = {
    "http":"http://127.0.0.1:4780",
}

def get_movies(id, piaofang):
    url = 'https://piaofang.maoyan.com/movie/{}'.format(id)
    headers = {
        'Host': 'piaofang.maoyan.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.90 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://piaofang.maoyan.com/mdb/rank',
        'Connection': 'keep-alive',
        'Cookie': '_lxsdk_cuid=17848ee5ea0c8-0b52b3b2c36967-5771031-144000-17848ee5ea1c8; _lxsdk=09A77C80887311EBB1D641DA3EE7CFD01EF60398D63A403DA981344067F35A2F; __mta=50099865.1616131306519.1616155894519.1616156263535.3; theme=moviepro; __mta=48632601.1616080953387.1616160109317.1616160330772.57',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    print(headers)
    res = requests.get(url,headers=headers)
    print(res.content)
    html = etree.HTML(res.text)
    title = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[1]/div[1]/p/span/text()')
    type = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div/p/text()')
    fangying = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div/p/span/text()')
    country = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[1]/div/p/text()')
    times = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[1]/div/p/span/text()')
    shangying_time = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[1]/div[2]/div[2]/div[1]/div/div[2]/a/span[1]/text()')
    fen = html.xpath('/html/body/div[2]/section[1]/div[1]/div[2]/div[2]/a/div[2]/div[2]/div/span[1]/text()')
    if len(title) > 0:
        title = title[0]
    else:
        title = ' '
    if len(type) > 0:
        type = type[0].replace('\n','').replace(' ','')
    else:
        type = ' '
    if len(fangying) > 0:
        fangying = fangying[0]
    else:
        fangying = ' '
    if len(country) > 0:
        country = country[0].replace('\n','').replace(' ','').replace('/','')
    else:
        country = ' '
    if len(times) > 0:
        times = times[0].replace(' ','')
    else:
        times = ' '
    if len(shangying_time) > 0:
        shangying_time = shangying_time[0]
    else:
        shangying_time = ' '
    if len(fen) > 0:
        fen = fen[0]
    else:
        fen = ' '
    piaofang = piaofang
    return (title,type,fangying,country,times,shangying_time,fen,piaofang)

def get_actors(id):
    url = "https://piaofang.maoyan.com/movie/{}/celebritylist".format(id)
    headers = {
        'Host': 'piaofang.maoyan.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:86.0) Gecko/20100101 Firefox/86.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://piaofang.maoyan.com/movie/{}'.format(id),
        'Connection': 'keep-alive',
        'Cookie': '_lxsdk_cuid=17848ee5ea0c8-0b52b3b2c36967-5771031-144000-17848ee5ea1c8; _lxsdk=09A77C80887311EBB1D641DA3EE7CFD01EF60398D63A403DA981344067F35A2F; _lxsdk_s=17849ad3719-3ff-4dd-77f||1; __mta=48632601.1616080953387.1616135751747.1616145379452.24; theme=moviepro',
        'Upgrade-Insecure-Requests': '1',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    res = requests.get(url=url,headers=headers)
    html = etree.HTML (res.text)
    daoyans = html.xpath("/html/body/div/div/div[1]/div/dl[1]/dd/div/div/div/div[2]/p[1]/text()")
    yanyuans = html.xpath('/html/body/div/div/div[1]/div/dl[2]/dd/div/div[1]/div[1]/div[2]/p[1]/text()')
    if len(daoyans) == 0:
        daoyans = [' ']
    if len (yanyuans) == 0:
       yanyuans = [' ']
    return ('/'.join(daoyans),'/'.join(yanyuans))
# print(res.text)


def get_company(id):
    url = "https://piaofang.maoyan.com/movie/{}/companylist".format (id)
    headers = {
        'Host': 'piaofang.maoyan.com',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:86.0) Gecko/20100101 Firefox/86.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2',
        'Accept-Encoding': 'gzip, deflate, br',
        'Referer': 'https://piaofang.maoyan.com/movie/{}'.format (id),
        'Connection': 'keep-alive',
        'Cookie': '__mta=48632601.1616080953387.1616156550165.1616156823356.44; _lxsdk_cuid=17848ee5ea0c8-0b52b3b2c36967-5771031-144000-17848ee5ea1c8; _lxsdk=09A77C80887311EBB1D641DA3EE7CFD01EF60398D63A403DA981344067F35A2F; __mta=50099865.1616131306519.1616155894519.1616156263535.3; _lxsdk_s=1784a65e193-a9a-811-90||14; __mta=48632601.1616080953387.1616156187664.1616156546133.43',
        'Upgrade-Insecure-Requests': '1',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache',
    }
    res = requests.get (url=url, headers=headers)
    html = etree.HTML (res.text)
    chupin = html.xpath ("/html/body/div[2]/div/div/div[1]/dl[1]/dd/div/div/div/div[2]/p/text()")
    lianhechupin = html.xpath ("/html/body/div[2]/div/div/div[1]/dl[2]/dd/div/div/div/div[2]/p/text()")
    faxing = html.xpath ('/html/body/div[2]/div/div/div[1]/dl[3]/dd/div/div/div/div[2]/p/text()')
    lianhefaxing = html.xpath ('/html/body/div[2]/div/div/div[1]/dl[4]/dd/div/div/div/div[2]/p/text()')
    qita = html.xpath ('/html/body/div[2]/div/div/div[1]/dl[5]/dd/div/div/div/div[2]/p/text()')
    if len(chupin) == 0:
        chupin =[' ']
    if len(lianhechupin) == 0:
        lianhechupin = [' ']
    if len(faxing) == 0:
        faxing = [' ']
    if len(lianhefaxing) == 0:
        lianhefaxing = [' ']
    if len(qita) == 0:
        qita = [' ']
    return ('/'.join(chupin),'/'.join(lianhechupin),'/'.join(faxing),'/'.join(lianhefaxing),'/'.join(qita))


# print(res.text)
import time
import os
if __name__ == '__main__':
    txt_lis = os.listdir('txt')
    for txt in txt_lis:
        workbook = xlsxwriter.Workbook('{}_movies.xls'.format(txt.replace('.txt','')))
        worksheet = workbook.add_worksheet('movies')
        cols = ['电影名称','电影类型','放映格式','制片国家和地区','时长','上映时间','评分','累计票房','导演','演员','出品公司','联合出品','放映公司','联合放映公司','其他','id']
        for i in range(16):
            worksheet.write(0,i,cols[i])
        with open('txt\{}.txt'.format(txt.replace('.txt','')),'r') as f:
            all = f.readlines()
            all = [i.replace('\n','').split(' ') for i in all]
            index = 1
            for i in all:
                print(i)
                try:
                    # title,type,fangying,country,times,shangying_time,fen,piaofang = get_movies(i[0],i[2])
                    # daoyan,yanyuan = get_actors(i[0])
                    # chupin,lianhechupin,fangying,lianhefangying,qita = get_company(i[0])
                    movie_data = get_movies(i[0],i[2])
                    actor_data = get_actors(i[0])
                    company_data = get_company(i[0])
                    all = list(movie_data) + list(actor_data) + list(company_data) + [i[0]]
                    for i in range(15):
                        worksheet.write (index, i, str(all[i]))
                    # with open('2011_movie.txt','a+') as f:
                    #     f.write('*'.join(all))
                    #     f.write('/n')
                    print(all)
                    index +=1
                except Exception as e:
                    print(e)
                    # with open('error.txt','a+',encoding='utf-8') as f:
                    #     f.write(i[0])
                    #     f.write('\n')
            workbook.close()

    #
    # cols = ['电影名称','电影类型','放映格式','制片国家和地区','时长','上映时间','评分','累计票房','导演','演员','出品公司','联合出品','放映公司','联合放映公司','其他']
    # with open ('all_movies.txt', 'r') as f:
    #     all = f.readlines ()
    #     all = [i.replace ('\n', '').split (' ') for i in all]
    # def process(i):
    #     try:
    #         # title,type,fangying,country,times,shangying_time,fen,piaofang = get_movies(i[0],i[2])
    #         # daoyan,yanyuan = get_actors(i[0])
    #         # chupin,lianhechupin,fangying,lianhefangying,qita = get_company(i[0])
    #         movie_data = get_movies(i[0],i[2])
    #         actor_data = get_actors(i[0])
    #         company_data = get_company(i[0])
    #         all = list(movie_data) + list(actor_data) + list(company_data)
    #         with open('result.txt','a+',encoding='utf-8') as f:
    #             f.write('*'.join(all))
    #             f.write('\n')
    #     except:
    #         with open('error.txt','a+',encoding='utf-8') as f:
    #             f.write(i[0])
    #             f.write('\n')
    # pools = ThreadPoolExecutor(max_workers=1)
    # pools.map(process,all)