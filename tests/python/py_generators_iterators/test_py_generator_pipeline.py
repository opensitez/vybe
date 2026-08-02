# vybe-test: python/py_generators_iterators/test_py_generator_pipeline
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def integers():
    n = 1
    while True:
        yield n
        n += 1

def squares(it):
    for n in it:
        yield n * n

def take(n, it):
    for _ in range(n):
        yield next(it)

pipeline = take(5, squares(integers()))
print(list(pipeline))
