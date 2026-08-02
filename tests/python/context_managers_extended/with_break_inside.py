# vybe-test: python/context_managers_extended/with_break_inside
# origin: languages/python/tests/python/test_context_managers_extended.rs

class CM:
 def __enter__(self):
  return self
 def __exit__(self, *a):
  pass
for i in range(2):
 with CM():
  if i:
   break
print('done')
