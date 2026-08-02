# vybe-test: python/builtins/builtin_callable
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

def f(): pass
b = callable(f)
