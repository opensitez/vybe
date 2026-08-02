# vybe-test: python/comprehension_walrus_spec/list_comp_if_chain_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [x for x in range(10) if x % 2 == 0 if x > 4]
