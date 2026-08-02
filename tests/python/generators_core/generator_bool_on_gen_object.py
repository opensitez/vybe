# vybe-test: python/generators_core/generator_bool_on_gen_object
# origin: languages/python/tests/python/test_generators_core.rs

g = (x for x in range(1))
print(bool(g))
