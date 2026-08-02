# vybe-test: python/context_managers_core/with_loop_repeated_entry
# origin: languages/python/tests/python/test_context_managers_core.rs

class CM:
 def __enter__(self):
  return 1
 def __exit__(self, *a):
  pass
total = 0
for _ in range(3):
 with CM() as v:
  total += v
print(total)
