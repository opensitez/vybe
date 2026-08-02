# vybe-test: python/closure_extended/closure_in_loop
# origin: languages/python/tests/python/test_closure_extended.rs

def make():
 out = []
 for i in range(3):
  out.append(lambda i=i: i * 2)
 return out
print(make()[2]())
