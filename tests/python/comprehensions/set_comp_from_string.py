# vybe-test: python/comprehensions/set_comp_from_string
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

s = {c for c in 'hello world' if c != ' '}
