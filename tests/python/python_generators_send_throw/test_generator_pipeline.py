# vybe-test: python/python_generators_send_throw/test_generator_pipeline
# origin: languages/python/tests/python/test_python_generators_send_throw.rs

def producer():
    for i in range(5):
        yield i

def doubler(gen):
    for x in gen:
        yield x * 2

def consumer(gen):
    return list(gen)

result = consumer(doubler(producer()))
print(result)
