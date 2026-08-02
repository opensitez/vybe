# vybe-test: python/py_generators_iterators/test_py_generator_basic_yield
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def count_up(n):
    for i in range(n):
        yield i

gen = count_up(4)
print(next(gen))
print(next(gen))
print(list(gen))  # exhaust remaining
