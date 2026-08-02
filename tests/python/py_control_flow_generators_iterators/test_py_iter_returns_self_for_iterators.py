# vybe-test: python/py_control_flow_generators_iterators/test_py_iter_returns_self_for_iterators
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

g = (x for x in range(3))
print(iter(g) is g)
