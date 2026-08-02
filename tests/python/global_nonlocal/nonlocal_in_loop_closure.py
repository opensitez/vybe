# vybe-test: python/global_nonlocal/nonlocal_in_loop_closure
# origin: languages/python/tests/python/test_global_nonlocal.rs

def outer():
 total = 0
 def add(n):
  nonlocal total
  total += n
 for i in range(3):
  add(i)
 return total
print(outer())
