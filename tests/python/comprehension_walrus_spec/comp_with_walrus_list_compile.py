# vybe-test: python/comprehension_walrus_spec/comp_with_walrus_list_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

xs = [(y := x * 2) for x in range(3)]
