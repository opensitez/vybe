# vybe-test: python/dict_comprehension/dict_comp_conditional_value
# origin: languages/python/tests/python/test_dict_comprehension.rs

{x: ('even' if x % 2 == 0 else 'odd') for x in range(3)}
