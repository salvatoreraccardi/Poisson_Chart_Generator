import math

def poisson(n, mi):
    value = math.exp(-mi) * ( math.pow(mi, n) / math.factorial(n) )
    return value