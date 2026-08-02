# vybe-test: python/error_handling/bare_raise_reraise
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

try:
    x = 1 / 0
except:
    print('logging')
    raise
