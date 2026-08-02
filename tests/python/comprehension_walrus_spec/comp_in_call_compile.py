# vybe-test: python/comprehension_walrus_spec/comp_in_call_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs
# vybe-test-mode: compile

print(sum(x * x for x in range(5)))
