# vybe-test: python/exceptions/try_except_else
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = 1
except:
    print("error")
else:
    print("no error")
