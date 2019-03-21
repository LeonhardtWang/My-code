

def fun(score_lis):
    if len(score_lis) == 1:
        return 1
    if len(score_lis) == 2:
        if score_lis[0] != score_lis[1]:
            return 3
        else:
            return 2

    res = {}
    for i in range(len(score_lis)-1):
        res[i] = 1
    res[-1] = 1
    is_djust = []
    i = -1
    while True:
        if score_lis[i-1] < score_lis[i]:
            if i -1 == -2:
                if res[len(score_lis)-2] >= res[i]:
                    res[i] = res[len(score_lis)-2] + 1
                    is_djust.append(1)
            else:
                if res[i-1] >= res[i]:
                    res[i] += 1
                    is_djust.append(1)
        if i+1 == len(score_lis) - 1:
            if not is_djust:
                break
            else:
                i = -1
                is_djust = []
                continue
        if score_lis[i] > score_lis[i+1]:
            if res[i] <= res[i+1]:
                res[i] = res[i+1] + 1
                is_djust.append(1)
        i += 1
    return sum(list(res.values()))

for i in [[1,2,3,3], [1,2]]:
    print(fun(i))


