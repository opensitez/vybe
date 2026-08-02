# vybe-test: python/syntax/args_variadic
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def f(*args):
    print(args)
f(1, 2, 3)
