# vybe-test: python/context_managers_core/with_else_not_valid_use_plain_after
# origin: languages/python/tests/python/test_context_managers_core.rs

class CM:
 def __enter__(self):
  return self
 def __exit__(self, *a):
  return False
try:
 with CM():
  pass
 print('after')
except:
 print('no')
