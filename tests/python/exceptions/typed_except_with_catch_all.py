# vybe-test: python/exceptions/typed_except_with_catch_all
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    x = 1
except ValueError:
    print("value error")
except:
    print("catch all")
