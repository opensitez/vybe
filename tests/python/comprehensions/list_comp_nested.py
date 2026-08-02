# vybe-test: python/comprehensions/list_comp_nested
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

flat = [x for row in [[1,2],[3,4]] for x in row]
