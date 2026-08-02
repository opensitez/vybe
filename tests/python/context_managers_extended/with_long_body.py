# vybe-test: python/context_managers_extended/with_long_body
# origin: languages/python/tests/python/test_context_managers_extended.rs

class CM:
 def __enter__(self):
  return 0
 def __exit__(self, *a):
  pass
with CM() as s:
 for i in range(3):
  s += i
 print(s)
