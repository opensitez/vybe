# vybe-test: python/functions/map_basic
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

result = list(map(lambda x: x * 2, [1, 2, 3]))
print(result)
