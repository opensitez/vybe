# vybe-test: python/py_concurrency/test_py_concurrent_futures_threadpool
# origin: languages/python/tests/python/test_py_concurrency.rs

from concurrent.futures import ThreadPoolExecutor

def compute(x):
    return x * x

with ThreadPoolExecutor(max_workers=4) as ex:
    futures = [ex.submit(compute, i) for i in range(5)]
    results = sorted(f.result() for f in futures)

print(results)
