# vybe-test: python/comprehension_walrus_spec/walrus_in_ifexpr_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

x = 1
y = (z := x + 1) if x else 0
