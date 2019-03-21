def fun(n):
    n = 1024 - n 
    def recursive(n):
        if n <= 4:
            return n
        elif 4 < n <= 16:
            return n // 4 + recursive(n - (n // 4) * 4)
        elif 16 < n <= 64:
            return n // 16 + recursive(n - (n //16) * 16)
        else:
            return n // 64 + recursive(n - (n // 64) * 64)
    res = recursive(n)
    return res


print(fun(200))