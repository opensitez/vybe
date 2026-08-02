# vybe-test: python/py_control_flow_generators_iterators/test_py_itertools_islice_range_slicing
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

from itertools import islice

def infinite_counter():
    n = 0
    while True:
        yield n
        n += 1

slice_out = list(islice(infinite_counter(), 5, 10))
print(slice_out)
