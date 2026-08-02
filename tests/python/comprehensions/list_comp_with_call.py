# vybe-test: python/comprehensions/list_comp_with_call
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

upper = [s.upper() for s in ['a', 'b', 'c']]
