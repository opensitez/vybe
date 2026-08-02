# vybe-test: python/py_generator_coroutines_yield/test_py_generator_send_bidirectional_data
# origin: languages/python/tests/python/test_py_generator_coroutines_yield.rs

def echo_coroutine():
    val = yield "ready"
    while True:
        val = yield f"echo: {val}"

co = echo_coroutine()
print(next(co))
print(co.send("first"))
print(co.send("second"))
