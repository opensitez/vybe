# vybe-test: python/context_managers_core/with_exit_returns_true_suppresses
# origin: languages/python/tests/python/test_context_managers_core.rs

class CM:
 def __enter__(self):
  return self
 def __exit__(self, exc, val, tb):
  return True
try:
 with CM():
  raise ValueError('x')
except ValueError:
 print('leaked')
print('after')
