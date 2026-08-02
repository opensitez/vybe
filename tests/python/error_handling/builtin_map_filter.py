# vybe-test: python/error_handling/builtin_map_filter
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

result = list(map(str, [1, 2, 3]))
result2 = list(filter(None, [0, 1, 2]))
