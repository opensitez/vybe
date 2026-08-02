# vybe-test: python/generators_core/generator_map_on_genexp
# origin: languages/python/tests/python/test_generators_core.rs

list(map(lambda x: x * 2, (i for i in range(3))))
