# vybe-test: python/comprehension_walrus_spec/set_comp_runtime_len
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

s = {x % 2 for x in range(6)}
print(len(s))
