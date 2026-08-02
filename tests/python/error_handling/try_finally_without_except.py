# vybe-test: python/error_handling/try_finally_without_except
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

try:
    f = open('x')
finally:
    f.close()
