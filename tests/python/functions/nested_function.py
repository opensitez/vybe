# vybe-test: python/functions/nested_function
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

def outer():
    def inner():
        return 42
    return inner()
