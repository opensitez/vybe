# vybe-test: python/error_handling/raise_from
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

try:
    pass
except Exception as e:
    raise RuntimeError('wrapped') from e
