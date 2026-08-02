# vybe-test: python/builtins/builtin_enumerate
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

for i, v in enumerate([1,2,3]):
    print(i, v)
