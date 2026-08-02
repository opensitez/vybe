# vybe-test: python/py_control_flow_generators_iterators/test_py_generator_expression_chaining
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

nums = range(10)
evens = (x for x in nums if x % 2 == 0)
squared = (x * x for x in evens)
print(list(squared))
