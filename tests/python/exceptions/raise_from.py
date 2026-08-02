# vybe-test: python/exceptions/raise_from
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = 1 / 0
except ZeroDivisionError as e:
    raise ValueError("invalid") from e
