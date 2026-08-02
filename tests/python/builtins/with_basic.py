# vybe-test: python/builtins/with_basic
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

with open('file.txt') as f:
    data = f.read()
