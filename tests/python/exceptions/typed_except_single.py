# vybe-test: python/exceptions/typed_except_single
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = int("abc")
except ValueError:
    print("bad value")
