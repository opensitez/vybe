# vybe-test: python/py_multiprocessing_pool_processes/test_py_multiprocessing_shared_value_and_array
# origin: languages/python/tests/python/test_py_multiprocessing_pool_processes.rs

import multiprocessing

def f(n, a):
    n.value = 3.14159
    for i in range(len(a)):
        a[i] = -a[i]

if __name__ == "__main__":
    num = multiprocessing.Value('d', 0.0)
    arr = multiprocessing.Array('i', range(5))

    p = multiprocessing.Process(target=f, args=(num, arr))
    p.start()
    p.join()

    print(round(num.value, 2))
    print(list(arr))
