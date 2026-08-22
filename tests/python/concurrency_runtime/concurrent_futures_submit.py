# vybe-test: python/concurrency_runtime/concurrent_futures_submit
# origin: languages/python/tests/python/test_concurrency_runtime.rs

import concurrent.futures
with concurrent.futures.ThreadPoolExecutor() as ex:
 ex.submit(lambda: 1)
