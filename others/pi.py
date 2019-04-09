import random


def RandomPI(N = 1000):
    inCircle = 0
    for i in range(N):
        if(random.random()**2 + random.random()**2 < 1):
            inCircle+=1
    print('The value of PI is ' + str(inCircle / N * 4))



def SeriesPI(N=1000):
    pi = 1
    n = 3
    neg = -1
    for i in range(N):
        pi+=neg*1/n
        neg*=-1
        n+=2
    print('The value of PI is {0:.10f}'.format(pi*4))


def SeriesPI2(N=10):
    pi = 1
    a = 1
    b = 3
    c = 1
    for i in range(N):
        c *= a/b
        pi+=c
        a+=1
        b+=2
    print('The value of PI is {0:.17f}'.format(pi*2))


if __name__ == '__main__':
    SeriesPI2(10)
    SeriesPI2(20)
    SeriesPI2(30)
    SeriesPI2(40)
    SeriesPI2(50)
    # SeriesPI(1000)

    # RandomPI(1000)