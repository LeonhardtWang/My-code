import pymongo
import datetime
import pandas as pd
import unicodedata
import math
import os

def dataframe(item_lis):
    col_name = []
    for item in item_lis:
        if len(item) > len(col_name):
            col_name = item
    col_name = list(col_name.keys())
    res = []
    for item in item_lis:
        item_val = []
        for key in col_name:
            if key in item:
                item_val.append(item[key])
            else:
                item_val.append('')
        res.append(item_val)
    return col_name, res

class PrintCMD:
    def __init__(self):
        pass

    def chr_width(self, c):
        if not c:return 0
        if (unicodedata.east_asian_width(c) in ('F','W','A')):
            return 2
        else:
            return 1

    def str_width(self, string):
        width = 0
        for s in string:
            width += self.chr_width(s)
        return width  

    def pri_beauty(self, row, max_indent, indent):
        max_indent = list(max_indent.values())
        fun_wid_chr = self.str_width
        while max([len(str(s)) for s in row]) != 0:
            print_s = ''
            for i, s in enumerate(row):
                if not isinstance(s, str):
                    if math.isnan(s):
                        s = ''
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

                if max_indent[i] < indent:
                    add_s = add_s + ' ' * (max_indent[i] - fun_wid_chr(add_s)) + '|'                  
                else:
                    add_s = add_s + ' ' * (indent - fun_wid_chr(add_s)) + '|'
                print_s += add_s
            print('|'+print_s)



class FrontEdge:
    def __init__(self):
        self.client = pymongo.MongoClient(host='58.220.2.26', port=27017)
        self.db_test = self.client['test']
        self.db_user = self.client['设备资产']

    def output_excel(self, collect_name, query_val, is_print=True, indent=15):
        collection = self.db_user[collect_name]
        results = list(collection.find())
        if not results:
            print("数据库中%s无数据！" % collect_name)
            return 
        
        print("开始查询，请稍等……")
        # 全查询功能
        serach_res = []
        for res in results:
            del res['_id']
            del res['ID']
            is_match = True
            for q in query_val.split():
                q_is_match = False
                for val in res.values():
                    if q in val:
                        is_match = is_match & (q in val)
                        q_is_match = True
                        break
                is_match = is_match & q_is_match

            if is_match:
                serach_res.append(res)
        
        if not serach_res:
            print("%s无任何符合查询条件的数据！" % collect_name)
            return 

        #df = pd.DataFrame(serach_res)      
        if is_print:
            print("查询结果如下：")
            self.print_info(serach_res, indent)
        else:
            # excel命名为时间
            time_now = str(datetime.datetime.now()).split('.')[0].split(' ')
            time_now = time_now[0] + '#'+ time_now[1]
            time_now = time_now.replace(':', '-')
            
            col_name, res = dataframe(serach_res)
            df = pd.DataFrame(res, columns=col_name)
            df.to_excel(time_now+'.xlsx', index=False)
            print("生成EXcel完成：%s" % (time_now + '.xlsx'))


    def print_info(self, serach_res, indent):
        printcmd = PrintCMD()
        
        col_name, res = dataframe(serach_res)

        max_indent = {}
        for row in res:
            for i in range(len(row)):
                if not isinstance(row[i], str):
                    if math.isnan(row[i]):
                        row[i] = ''
                if col_name[i] not in max_indent:
                    max_indent[col_name[i]] = printcmd.str_width(row[i])
                else:
                    if max_indent[col_name[i]] < printcmd.str_width(row[i]):
                        max_indent[col_name[i]] = printcmd.str_width(row[i])
        for col in col_name:
            if max_indent[col] < printcmd.str_width(col):
                max_indent[col] = printcmd.str_width(col)

        print_len = 1

        for val in max_indent.values():
            if val > indent:
                print_len += indent
                print_len += 1
            else:
                print_len += val
                print_len += 1

        print('-' * print_len)     
        printcmd.pri_beauty(col_name, max_indent, indent)
        print('-' * print_len)          
        for row in res:
            printcmd.pri_beauty(row, max_indent, indent)
            print('-' * print_len)


if __name__ == '__main__':
    print('欢迎进入工单管理系统1.0'.center(30, '☆'))
    print('\n\n')
    while True:
        query = input("\n是否继续，y or n？")
        if query == 'n':
            print("程序终止！")
            os._exit(0)
        if query not in ['y', 'n']:
            print("输入错误！")
            continue
            
        print('↓' * 30)    
        print("请选择查询哪个表（输入序号即可）：")
        collect_lis = ["设备资产线上表", "设备资产库存表", "设备资产变更表"]
        for i, clt in enumerate(collect_lis):
            print(i+1, clt)
        while True:
            c_name_idx = input("请输入：")
            if c_name_idx not in [str(i+1) for i in range(len(collect_lis))]:
                print("输入错误！")
            else:
                break
        collect_name = collect_lis[int(c_name_idx)-1]
        print('↓' * 30)

        query_val = input("请输入查询值（用空格隔开可查询多个值）：")
        while True:
            is_print = input("是否生成Excel，y or n?")
            if is_print == 'y':
                is_print = False
                break
            elif is_print == 'n':
                is_print = True
                break
            else:
                print("输入错误！")
        print('↓' * 30)
        

        front_edge = FrontEdge()
        front_edge.output_excel(collect_name, query_val, is_print=is_print, indent=25)


