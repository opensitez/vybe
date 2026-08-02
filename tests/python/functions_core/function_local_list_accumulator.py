# vybe-test: python/functions_core/function_local_list_accumulator
# origin: languages/python/tests/python/test_functions_core.rs

def build():
 out = []
 for i in range(3):
  out.append(i)
 return out
print(build())
