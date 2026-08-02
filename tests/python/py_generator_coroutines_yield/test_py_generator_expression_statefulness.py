# vybe-test: python/py_generator_coroutines_yield/test_py_generator_expression_statefulness
# origin: languages/python/tests/python/test_py_generator_coroutines_yield.rs

squares = (x * x for x in range(5))
print(next(squares))
print(next(squares))
print(list(squares))  # consumes remainder
print(list(squares))  # empty now
