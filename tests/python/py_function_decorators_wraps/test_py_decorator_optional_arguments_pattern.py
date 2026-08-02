# vybe-test: python/py_function_decorators_wraps/test_py_decorator_optional_arguments_pattern
# origin: languages/python/tests/python/test_py_function_decorators_wraps.rs

from functools import wraps, partial

def smart_decorator(func=None, *, prefix="LOG"):
    if func is None:
        return partial(smart_decorator, prefix=prefix)

    @wraps(func)
    def wrapper(*args, **kwargs):
        print(f"[{prefix}] Calling {func.__name__}")
        return func(*args, **kwargs)
    return wrapper

@smart_decorator
def f1(): return "f1"

@smart_decorator(prefix="CUSTOM")
def f2(): return "f2"

f1()
f2()
