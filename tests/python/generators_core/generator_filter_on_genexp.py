# vybe-test: python/generators_core/generator_filter_on_genexp
# origin: languages/python/tests/python/test_generators_core.rs

list(filter(lambda x: x > 1, (i for i in range(4))))
