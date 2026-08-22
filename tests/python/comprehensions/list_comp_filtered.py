# vybe-test: python/comprehensions/list_comp_filtered
# origin: languages/python/tests/python/test_comprehensions.rs

result = [x for x in range(10) if x % 2 == 0]
