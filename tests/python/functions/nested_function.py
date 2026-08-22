# vybe-test: python/functions/nested_function
# origin: languages/python/tests/python/test_functions.rs

def outer():
    def inner():
        return 42
    return inner()
