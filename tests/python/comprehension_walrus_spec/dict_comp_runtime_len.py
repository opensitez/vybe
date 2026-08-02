# vybe-test: python/comprehension_walrus_spec/dict_comp_runtime_len
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

d = {x: x * x for x in range(4)}
print(len(d))
