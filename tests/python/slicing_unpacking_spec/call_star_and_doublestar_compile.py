# vybe-test: python/slicing_unpacking_spec/call_star_and_doublestar_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs

def f(a, b, c):
    pass
args = [1]
kw = {'b': 2, 'c': 3}
f(*args, **kw)
