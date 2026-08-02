# vybe-test: python/exceptions/nested_try
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    try:
        x = 1 / 0
    except ZeroDivisionError:
        print("inner")
        raise ValueError("converted")
except ValueError:
    print("outer")
