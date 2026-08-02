# vybe-test: python/exceptions/try_except_finally
# origin: languages/python/tests/python/test_exceptions.rs
# vybe-test-mode: compile

try:
    f = open("file.txt")
except:
    print("error")
finally:
    print("cleanup")
