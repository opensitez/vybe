# vybe-test: python/syntax/call_double_star_unpack
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def f(a, b):
    pass
kw = {'a': 1, 'b': 2}
f(**kw)
