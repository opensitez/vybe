# vybe-test: python/dict_comprehension/dict_comp_any_value_truthy
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {i: bool(i) for i in range(3)}
print(any(d.values()))
