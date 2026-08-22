# vybe-test: python/error_handling/except_tuple_of_types
# origin: languages/python/tests/python/test_error_handling.rs

try:
    pass
except (ValueError, TypeError):
    pass
