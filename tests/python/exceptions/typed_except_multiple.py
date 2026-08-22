# vybe-test: python/exceptions/typed_except_multiple
# origin: languages/python/tests/python/test_exceptions.rs

try:
    x = 1 / 0
except ValueError:
    print("value error")
except TypeError:
    print("type error")
except ZeroDivisionError:
    print("division by zero")
