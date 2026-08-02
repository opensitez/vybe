# vybe-test: python/syntax/nonlocal_statement
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def outer():
    x = 10
    def inner():
        nonlocal x
        x = 20
    inner()
