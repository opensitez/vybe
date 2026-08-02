# vybe-test: python/context_managers_core/with_break_inside_block
# origin: languages/python/tests/python/test_context_managers_core.rs

class CM:
 def __enter__(self):
  return self
 def __exit__(self, *a):
  pass
for _ in range(2):
 with CM():
  print('x')
  break
