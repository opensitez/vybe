# vybe-test: python/py_generators_iterators/test_py_generator_expression
# origin: languages/python/tests/python/test_py_generators_iterators.rs

squares = (x ** 2 for x in range(5))
print(type(squares).__name__)
print(next(squares))
print(sum(squares))  # 1+4+9+16 = 30
