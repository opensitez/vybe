# vybe-test: python/dict_comprehension/dict_comp_list_value_mutation_independent
# origin: languages/python/tests/python/test_dict_comprehension.rs

d = {i: [] for i in range(2)}
d[0].append(1)
print(d[1])
