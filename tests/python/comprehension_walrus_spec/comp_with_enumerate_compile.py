# vybe-test: python/comprehension_walrus_spec/comp_with_enumerate_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

pairs = [(i, x) for i, x in enumerate(['a', 'b'])]
