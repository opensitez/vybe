# vybe-test: python/py_function_decorators_wraps/test_py_decorator_with_arguments_factory
# origin: languages/python/tests/python/test_py_function_decorators_wraps.rs

from functools import wraps

def repeat(num_times):
    def decorator_repeat(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for _ in range(num_times):
                result = func(*args, **kwargs)
            return result
        return wrapper
    return decorator_repeat

log = []

@repeat(num_times=3)
def ping():
    log.append("pong")

ping()
print(log)
