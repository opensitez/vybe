# vybe-test: python/py_control_flow_generators_iterators/test_py_itertools_chain_generator_sequences
# origin: languages/python/tests/python/test_py_control_flow_generators_iterators.rs

from itertools import chain

g1 = (x for x in range(2))
g2 = (x * 10 for x in range(1, 3))
chained = list(chain(g1, g2))
print(chained)
