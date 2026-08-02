# vybe-test: python/functions/filter_basic
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

result = list(filter(lambda x: x > 0, [-1, 0, 1, 2]))
print(result)
