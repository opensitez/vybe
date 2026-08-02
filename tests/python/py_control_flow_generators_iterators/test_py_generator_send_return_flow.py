# vybe-test: python/py_control_flow_generators_iterators/test_py_generator_send_return_flow
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

def accumulator():
    total = 0
    while True:
        val = yield total
        if val is None:
            break
        total += val

acc = accumulator()
print(next(acc))  # prime
print(acc.send(10))
print(acc.send(20))
