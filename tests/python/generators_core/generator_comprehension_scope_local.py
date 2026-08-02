# vybe-test: python/generators_core/generator_comprehension_scope_local
# origin: languages/python/tests/python/test_generators_core.rs

out = (x for x in range(2))
print(list(out))
