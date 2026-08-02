# vybe-test: python/builtins/builtin_dir
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

x = {'a': 1, 'b': 2}
keys = dir(x)
