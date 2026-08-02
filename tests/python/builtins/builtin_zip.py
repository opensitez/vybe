# vybe-test: python/builtins/builtin_zip
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

for a, b in zip([1,2], [3,4]):
    print(a, b)
