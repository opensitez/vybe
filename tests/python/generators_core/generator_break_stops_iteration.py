# vybe-test: python/generators_core/generator_break_stops_iteration
# origin: languages/python/tests/python/test_generators_core.rs

def g():
 for i in range(5):
  yield i
count = 0
for _ in g():
 count += 1
 if count == 2:
  break
print(count)
