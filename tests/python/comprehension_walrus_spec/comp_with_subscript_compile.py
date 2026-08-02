# vybe-test: python/comprehension_walrus_spec/comp_with_subscript_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [row[0] for row in [[1, 2], [3, 4]]]
