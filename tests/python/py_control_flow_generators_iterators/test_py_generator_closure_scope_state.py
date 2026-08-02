# vybe-test: python/py_control_flow_generators_iterators/test_py_generator_closure_scope_state
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

def make_gen(factor):
    return (x * factor for x in range(3))

g = make_gen(10)
print(list(g))
