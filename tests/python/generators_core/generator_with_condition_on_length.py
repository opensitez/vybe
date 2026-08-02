# vybe-test: python/generators_core/generator_with_condition_on_length
# origin: languages/python/tests/python/test_generators_core.rs

print(list(s for s in ['a', 'bb'] if len(s) == 1))
