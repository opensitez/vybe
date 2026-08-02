# vybe-test: python/comprehension_walrus_spec/walrus_nested_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

result = [z for x in range(3) if (y := x + 1) and (z := y + 1)]
