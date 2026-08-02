# vybe-test: python/generators_core/generator_len_not_supported
# origin: languages/python/tests/python/test_generators_core.rs

g = (x for x in range(3))
try:
 len(g)
except TypeError:
 print('no')
