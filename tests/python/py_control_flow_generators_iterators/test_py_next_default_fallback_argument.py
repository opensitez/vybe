# vybe-test: python/py_control_flow_generators_iterators/test_py_next_default_fallback_argument
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

g = (x for x in [1, 2])
print(next(g, "end"))
print(next(g, "end"))
print(next(g, "end"))
