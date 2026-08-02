# vybe-test: python/syntax/decorator_complex_expression
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

@module.sub.decorator(arg1, key=val)
def func():
    pass
