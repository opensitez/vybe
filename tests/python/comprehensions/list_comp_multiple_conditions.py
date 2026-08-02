# vybe-test: python/comprehensions/list_comp_multiple_conditions
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

result = [x for x in range(100) if x % 2 == 0 if x % 3 == 0]
