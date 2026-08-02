# vybe-test: python/comprehension_walrus_spec/comp_with_zip_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

pairs = [(a, b) for a, b in zip([1, 2], [3, 4])]
