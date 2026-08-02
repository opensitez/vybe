# vybe-test: python/set_comprehension/set_comp_conditional_expression
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({('even' if x % 2 == 0 else 'odd') for x in range(3)})
