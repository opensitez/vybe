# vybe-test: python/syntax/with_nested
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

with open('a') as f:
    with open('b') as g:
        pass
