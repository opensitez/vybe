# vybe-test: python/dict_comprehension/dict_comp_sum_values
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {i: i for i in range(4)}
print(sum(d.values()))
