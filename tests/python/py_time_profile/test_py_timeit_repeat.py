# vybe-test: python/py_time_profile/test_py_timeit_repeat
# origin: languages/python/tests/python/test_py_time_profile.rs

import timeit

runs = timeit.repeat("sorted([3, 1, 4, 1, 5])", number=100, repeat=3)
print(len(runs))
print(all(r > 0 for r in runs))
