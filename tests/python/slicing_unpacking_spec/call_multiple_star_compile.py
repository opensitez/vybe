# vybe-test: python/slicing_unpacking_spec/call_multiple_star_compile
# origin: languages/python/tests/python/test_slicing_unpacking_spec.rs
# vybe-test-mode: compile

def f(a, b, c, d):
    pass
left = [1, 2]
right = [3, 4]
f(*left, *right)
