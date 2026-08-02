# vybe-test: python/py_closures_hof/test_py_closure_loop_capture_gotcha
# origin: languages/python/tests/python/test_py_closures_hof.rs

# Late binding gotcha
fns_bad = [lambda x: x + i for i in range(4)]
print(fns_bad[0](0))  # all capture same i=3

# Fixed with default argument
fns_good = [lambda x, i=i: x + i for i in range(4)]
print([f(0) for f in fns_good])
