n1 = 3
n2 = [8, 8, 8]
n3 = [2, 2, 2]
res = 0

for i in range(n1 * 2):
    if i % 2 == 0:
        l = n2; r = n3
    else:
        l = n3; r = n2
    if l and r:
        if max(l) < max(r):
            del r[r.index(max(r))]
        else:
            if i % 2 == 0:
                res += max(l)
            else:
                res -= max(l)
            del l[l.index(max(l))]
    elif not l and r:
        del r[r.index(max(r))]
    else:
        if i % 2 == 0:
            res += max(l)
        else:
            res -= max(l)
        del l[l.index(max(l))]

print(res)