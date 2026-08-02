# vybe-test: python/generators_core/generator_state_persists_between_yields
# origin: languages/python/tests/python/test_generators_core.rs

def g():
 x = 0
 while x < 3:
  yield x
  x += 1
print(list(g()))
