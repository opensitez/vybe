# vybe-test: python/syntax/stacked_decorators
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

@decorator1
@decorator2
def func():
    pass
