# vybe-test: python/py_generators_iterators/test_py_generator_yield_from_flattens_nested
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def flatten(nested):
    for item in nested:
        if isinstance(item, list):
            yield from flatten(item)
        else:
            yield item

data = [1, [2, [3, 4], 5], [6, 7]]
print(list(flatten(data)))
