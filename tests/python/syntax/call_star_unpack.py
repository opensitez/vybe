# vybe-test: python/syntax/call_star_unpack
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def f(a, b, c):
    pass
args = [1, 2, 3]
f(*args)
