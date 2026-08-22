# vybe-test: python/syntax/kwargs_variadic
# origin: languages/python/tests/python/test_syntax.rs

def f(**kwargs):
    print(kwargs)
f(a=1, b=2)
