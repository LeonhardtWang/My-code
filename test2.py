test = ['helllooaajj', 'woooooow']


def fun1(string):
    str_lis = list(string)
    del_index = []
    for i in range(len(str_lis)):
        if i + 2 >= len(str_lis):break
        if str_lis[i] == str_lis[i+1] == str_lis[i+2]:
            del_index.append(i)
    res = ''
    for i in range(len(str_lis)):
        if i not in del_index:
            res += str_lis[i]
    return res

def fun2(string):
    str_lis = list(string)
    del_index = []
    i = 0
    while i < len(str_lis):
        if i + 3 >= len(str_lis):break
        if str_lis[i] == str_lis[i+1] and str_lis[i+2] == str_lis[i+3]:
            del_index.append(i+2)
            i += 3
            continue
        i += 1
    res = ''
    for i in range(len(str_lis)):
        if i not in del_index:
            res += str_lis[i]
    return res

print(fun2(fun1(test[0])))
