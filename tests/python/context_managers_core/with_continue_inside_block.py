# vybe-test: python/context_managers_core/with_continue_inside_block
# origin: languages/python/tests/python/test_context_managers_core.rs

class CM:
 def __enter__(self):
  return self
 def __exit__(self, *a):
  pass
out = []
for i in range(3):
 with CM():
  if i == 1:
   continue
 out.append(i)
print(out)
