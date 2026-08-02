# vybe-test: python/comprehension_walrus_spec/comp_with_walrus_set_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

s = {(y := x * 2) for x in range(3)}
