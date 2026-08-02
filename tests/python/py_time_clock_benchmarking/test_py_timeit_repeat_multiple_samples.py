# vybe-test: python/py_time_clock_benchmarking/test_py_timeit_repeat_multiple_samples
# origin: languages/python/tests/python/test_py_time_clock_benchmarking.rs

import timeit

samples = timeit.repeat("sorted([5, 2, 8, 1])", repeat=3, number=50)
print(len(samples))
print(all(isinstance(s, float) and s > 0 for s in samples))
