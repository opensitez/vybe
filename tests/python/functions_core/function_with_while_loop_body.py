# vybe-test: python/functions_core/function_with_while_loop_body
# origin: languages/python/tests/python/test_functions_core.rs

def first_gt(threshold):
 n = 0
 while n <= threshold:
  n += 1
 return n
print(first_gt(3))
