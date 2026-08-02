# vybe-test: python/py_generators_iterators/test_py_generator_stateful_transformation
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def running_average():
    total = 0
    count = 0
    while True:
        val = yield (total / count) if count else 0
        if val is not None:
            total += val
            count += 1

g = running_average()
next(g)
g.send(10)
g.send(20)
result = g.send(30)
print(result)
