# vybe-test: python/operators_arithmetic/arithmetic_div_by_zero_raises
# origin: languages/python/tests/python/test_operators_arithmetic.rs

try:
 print(1/0)
except ZeroDivisionError:
 print('z')
