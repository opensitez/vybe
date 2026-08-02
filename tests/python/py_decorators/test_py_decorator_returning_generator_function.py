# vybe-test: python/py_decorators/test_py_decorator_returning_generator_function
# origin: languages/python/tests/python/test_py_decorators.rs

def multiply_yields(factor):
    def decorator(gen_func):
        def wrapper(*args, **kwargs):
            for val in gen_func(*args, **kwargs):
                yield val * factor
        return wrapper
    return decorator

@multiply_yields(10)
def generate_numbers():
    yield 1
    yield 2
    yield 3

print(list(generate_numbers()))
