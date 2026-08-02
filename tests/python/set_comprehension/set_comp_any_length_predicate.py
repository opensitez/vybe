# vybe-test: python/set_comprehension/set_comp_any_length_predicate
# origin: languages/python/tests/python/test_set_comprehension.rs

s = {len(w) for w in ['a', 'bb', 'ccc']}
print(len(s))
