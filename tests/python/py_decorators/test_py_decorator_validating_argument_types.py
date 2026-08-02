# vybe-test: python/py_decorators/test_py_decorator_validating_argument_types
# origin: languages/python/tests/python/test_py_decorators.rs

def enforce_types(*expected_types):
    def decorator(func):
        def wrapper(*args):
            for arg, expected in zip(args, expected_types):
                if not isinstance(arg, expected):
                    raise TypeError(f"Expected {expected.__name__}, got {type(arg).__name__}")
            return func(*args)
        return wrapper
    return decorator

@enforce_types(int, str)
def repeat_str(count, text):
    return text * count

print(repeat_str(3, "Hi"))
try:
    repeat_str("invalid", "Hi")
except TypeError as e:
    print("TypeError caught")
