# vybe-test: python/functions/list_comprehension
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

squares = [x * x for x in range(10)]
print(squares)
