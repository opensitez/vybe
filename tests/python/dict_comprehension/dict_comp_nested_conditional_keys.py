# vybe-test: python/dict_comprehension/dict_comp_nested_conditional_keys
# origin: languages/python/tests/python/test_dict_comprehension.rs

{('pos' if x > 0 else 'neg'): x for x in [-1, 1]}
