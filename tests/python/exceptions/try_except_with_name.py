# vybe-test: python/exceptions/try_except_with_name
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = 1 / 0
except Exception as e:
    print(e)
