# vybe-test: python/python_generators_send_throw/test_generator_send_priming
# origin: languages/python/tests/python/test_python_generators_send_throw.rs

def echo():
    while True:
        received = yield
        print(f"got: {received}")

g = echo()
next(g)  # prime
g.send("hello")
g.send("world")
g.close()
