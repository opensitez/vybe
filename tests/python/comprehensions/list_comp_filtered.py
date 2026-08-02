# vybe-test: python/comprehensions/list_comp_filtered
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

result = [x for x in range(10) if x % 2 == 0]
