# vybe-test: python/set_comprehension/set_comp_union_with_literal
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in range(2)} | {2, 3})
