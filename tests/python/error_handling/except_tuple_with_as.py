# vybe-test: python/error_handling/except_tuple_with_as
# origin: languages/python/tests/python/test_error_handling.rs

try:
    pass
except (ValueError, TypeError) as e:
    print(e)
