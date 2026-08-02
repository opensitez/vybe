# vybe-test: python/for_while_extended/for_with_inside
# origin: languages/python/tests/python/test_for_while_extended.rs

class CM:
 def __enter__(self):
  return 1
 def __exit__(self, *a):
  pass
for _ in range(1):
 with CM() as v:
  print(v)
