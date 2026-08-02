# vybe-test: python/syntax/with_multiple_managers
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

with open('a') as f1, open('b') as f2:
    pass
