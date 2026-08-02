# vybe-test: python/comprehension_walrus_spec/comp_with_unpack_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [a + b for a, b in [(1, 2), (3, 4)]]
