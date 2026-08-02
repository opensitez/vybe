# vybe-test: python/slicing_extended/slice_zero_step_error
# origin: languages/python/tests/python/test_slicing_extended.rs

a = [1, 2, 3]
try:
 a[::0]
 print('ok')
except ValueError:
 print('err')
