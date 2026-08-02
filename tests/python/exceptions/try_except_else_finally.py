# vybe-test: python/exceptions/try_except_else_finally
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = 1
except ValueError:
    print("value error")
else:
    print("success")
finally:
    print("done")
