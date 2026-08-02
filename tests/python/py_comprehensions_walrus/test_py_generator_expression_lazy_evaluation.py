# vybe-test: python/py_comprehensions_walrus/test_py_generator_expression_lazy_evaluation
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

import sys

numbers = range(10 ** 6)
gen_expr = (x ** 2 for x in numbers if x % 2 == 0)
list_comp = [x ** 2 for x in range(10) if x % 2 == 0]

gen_size = sys.getsizeof(gen_expr)
list_size = sys.getsizeof(list_comp)
print(gen_size < list_size)  # generator takes less memory
print(next(gen_expr))
print(list_comp)
