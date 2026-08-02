# vybe-test: python/py_generator_coroutines_yield/test_py_generator_accumulator_coroutine
# origin: languages/python/tests/python/test_py_generator_coroutines_yield.rs

def running_total():
    total = 0
    while True:
        val = yield total
        if val is None:
            break
        total += val

t = running_total()
next(t)  # prime
print(t.send(10))
print(t.send(20))
print(t.send(30))
