# vybe-test: python/py_generators_iterators/test_py_generator_send_value
# origin: languages/python/tests/python/test_py_generators_iterators.rs

def accumulator():
    total = 0
    while True:
        value = yield total
        if value is None:
            break
        total += value

gen = accumulator()
next(gen)  # prime the generator
print(gen.send(10))
print(gen.send(20))
print(gen.send(5))
