# vybe-test: python/list_comprehension/list_comp_conditional_expression_inside
# origin: languages/python/tests/python/test_list_comprehension.rs

[('even' if x % 2 == 0 else 'odd') for x in range(3)]
