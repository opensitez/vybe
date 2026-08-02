# vybe-test: python/python_generators_send_throw/test_generator_send_value
# origin: languages/python/tests/python/test_python_generators_send_throw.rs

def accumulator():
    total = 0
    while True:
        value = yield total
        if value is None:
            break
        total += value

g = accumulator()
next(g)  # prime
print(g.send(10))
print(g.send(20))
print(g.send(5))
