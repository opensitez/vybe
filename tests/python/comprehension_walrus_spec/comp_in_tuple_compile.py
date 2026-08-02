# vybe-test: python/comprehension_walrus_spec/comp_in_tuple_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

t = tuple(x for x in range(3))
