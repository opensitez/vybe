# vybe-test: python/functions_core/function_with_for_loop_sum
# origin: languages/python/tests/python/test_functions_core.rs

def sum_range(n):
 total = 0
 for i in range(n):
  total += i
 return total
print(sum_range(4))
