# vybe-test: python/py_time_clock_benchmarking/test_py_timeit_timer_class_with_callable
# origin: languages/python/tests/python/test_py_time_clock_benchmarking.rs

import timeit

def benchmark_target():
    return [i * 2 for i in range(50)]

timer = timeit.Timer(benchmark_target)
elapsed = timer.timeit(number=100)
print(elapsed > 0)
