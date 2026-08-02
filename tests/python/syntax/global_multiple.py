# vybe-test: python/syntax/global_multiple
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

def f():
    global a, b, c
    a = 1
