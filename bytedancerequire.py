import requests
import pymongo
import json
import pandas as pd
import random
import math
import unicodedata

def craw_byte_apply_info():
    '''
    爬取字节跳动招聘信息入数据库
    '''
    headers = {
        'accept': 'application/json, text/plain, */*',
        'accept-encoding': 'gzip, deflate, br',
        'accept-language': 'zh-CN,zh;q=0.9',
        'referer': 'https://job.bytedance.com/society',
        'tea-uid': '6660379483738949127',
        'user-agent': 'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 \
            (KHTML, like Gecko) Chrome/65.0.3325.181 Mobile Safari/537.36'
    }
    '''
    proxy = [
        {
            'http':'http://175.161.29.145',
            'https':'http://175.161.29.145'
        },
        {
            'http':'http://47.96.148.248',
            'https':'http://47.96.148.248'
        },
        {
            'http':'http://115.231.5.230',
            'https':'http://115.231.5.230'
        }
    ]
    '''
    url = 'https://job.bytedance.com/api/recruitment/position/list/?type=1&limit=10&offset=0'
    res = requests.get(url, headers=headers)
    js = json.loads(res.text)

    page_num = int(js['count'])
    client = pymongo.MongoClient(host='localhost', port=27017, connect=False)
    client.drop_database('bytedance')
    db = client['bytedance']
    collections = db['positions']
    print("招聘官网：%s" % 'https://job.bytedance.com/intern?summary=&city=&q1=&position_type=')
    print("正在爬取字节跳动招聘信息，请稍等…………")
    for i in range(0, page_num, 10):
        get_url = url[:-1] + str(i)
        try:
            res = requests.get(get_url,headers=headers)
            js = json.loads(res.text)
            positions = js['positions']
        except:
            print("链接提取信息有误：%s" % get_url)
            continue
        for pos in positions:
            item = {}
            if pos['sub_name']:
                item['岗位'] = pos['name'] + str('-') + pos['sub_name']
            else:
                item['岗位'] = pos['name']
            item['职位类型'] = pos['summary']
            item['工作地点'] = pos['city']
            item['工作要求'] = pos['requirement'].strip().replace('\n', ' ')
            item['岗位描述'] = pos['description'].strip().replace('\n', ' ')
            item['发布日期'] = pos['create_time']
            print("\r当前速度：{:.2f}%".format((i//10+1)*100/(math.ceil(page_num/10))), end='')
            collections.insert_one(item)
    print("\n字节跳动招聘信息更新完成")

# 以下代码弃用
'''
def get_str_width(string):
    #用于得到中英文混合字符串的真正长度
    widths = [
      (126, 1), (159, 0), (687, 1), (710, 0), (711, 1),
      (727, 0), (733, 1), (879, 0), (1154, 1), (1161, 0),
      (4347, 1), (4447, 2), (7467, 1), (7521, 0), (8369, 1),
      (8426, 0), (9000, 1), (9002, 2), (11021, 1), (12350, 2),
      (12351, 1), (12438, 2), (12442, 0), (19893, 2), (19967, 1),
      (55203, 2), (63743, 1), (64106, 2), (65039, 1), (65059, 0),
      (65131, 2), (65279, 1), (65376, 2), (65500, 1), (65510, 2),
      (120831, 1), (262141, 2), (1114109, 1),
    ]
    width = 0
    for each in string:
        if ord(each) == 0xe or ord(each) == 0xf:
            each_width = 0
            continue
        elif ord(each) <= 1114109:
            for num, wid in widths:
                if ord(each) <= num:
                    each_width = wid
                    width += each_width
                    break
            continue

        else:
            each_width = 1
        width += each_width

    return width

def print_cmd(row):
    #结构化print，不用担心有的值过长
    #参数：row，类型为可迭代对象,如list,array等
    k = 0
    max_len = max([len(i) for i in row])
    print('-'*(30 * len(row) + 3))
    while k <= max_len:
        print_s = ''
        for s in row:
            sli = s[k:k+15]
            if get_str_width(sli) < 30:
                sli += ' ' * (30 - get_str_width(sli))
            sli += '|'
            print_s += sli
        print(print_s)
        k += 15
'''
def chr_width(c):
    if not c:return 0
    if (unicodedata.east_asian_width(c) in ('F','W','A')):
        return 2
    else:
        return 1

def str_width(string):
    width = 0
    for s in string:
        width += chr_width(s)
    return width  

def pri_beauty(row, indent, fun_wid_chr):
    print('-' * ((indent+1) * len(row)))
    while max([len(str(s)) for s in row]) != 0:
        print_s = ''
        for i, s in enumerate(row):
            s = str(s)
            add_s = ''
            k = 0
            while indent - fun_wid_chr(add_s) >= 0:
                add_s += s[k:k+1]
                if not s[k:]:break
                k += 1
                if indent - fun_wid_chr(add_s) < 0:
                    k -= 1
                    add_s = add_s[:-1]
                    break
            row[i] = s[k:]
            add_s = add_s + ' ' * (indent - fun_wid_chr(add_s)) + '|'
            print_s += add_s
        print(print_s)

def find(condition):
    '''
    参数：
    condition:查询条件，dict类型
    '''
    client = pymongo.MongoClient(host='localhost', port=27017)
    db = client['bytedance']
    collections = db['positions']
    results = collections.find(condition)
    res = list(results)
    res_df = pd.DataFrame(res)

    pri_beauty(list(res_df.columns[1:]), 30, str_width)
    for _, row in res_df.iterrows():
        del row['_id']
        pri_beauty(list(row), 30, str_width)

if __name__ == "__main__":
    #craw_byte_apply_info()
    key_find = ['岗位', '职位类型', '工作地点', '发布日期']
    while True:
        cmd = input("\n请输入命令：\n可用命令：f-查询, e-结束\n")
        if cmd == 'f':
            print('可查询字段如下：')
            for s in key_find:print(s, end='   ')
            print("")
            k_v = {}
            while True:
                key, value = input("\n请输入查询字段及内容（用空格隔开）：\n").split(' ')
                k_v.update({key:{'$regex':value}})
                y_n = input("是否要继续输入，是：y，否：n\n")
                if y_n == 'n':
                    print("开始查询！")
                    break
            find(k_v)
        elif cmd == 'e':
            print("程序终止")
            break
        else:
            print("无效命令，请重新输入")