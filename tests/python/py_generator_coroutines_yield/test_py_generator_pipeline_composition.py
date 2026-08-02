# vybe-test: python/py_generator_coroutines_yield/test_py_generator_pipeline_composition
# origin: languages/python/tests/python/test_py_generator_coroutines_yield.rs

def numbers(n):
    for i in range(n):
        yield i

def evens(seq):
    for x in seq:
        if x % 2 == 0:
            yield x

def doubled(seq):
    for x in seq:
        yield x * 2

pipeline = doubled(evens(numbers(10)))
print(list(pipeline))
