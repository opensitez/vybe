# vybe-test: python/builtins/dict_setdefault
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

d = {}
v = d.setdefault('key', 42)
