# vybe-test: python/functions/list_comp_with_filter
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

evens = [x for x in range(20) if x % 2 == 0]
