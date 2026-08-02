# vybe-test: python/comprehension_walrus_spec/comp_with_reversed_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [x for x in reversed([1, 2, 3])]
