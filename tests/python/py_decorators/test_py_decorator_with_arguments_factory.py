# vybe-test: python/py_decorators/test_py_decorator_with_arguments_factory
# origin: languages/python/tests/python/test_py_decorators.rs

def repeat(num_times):
    def decorator(func):
        def wrapper(*args, **kwargs):
            for _ in range(num_times):
                func(*args, **kwargs)
        return wrapper
    return decorator

@repeat(num_times=3)
def greet():
    print("Hello!")

greet()
