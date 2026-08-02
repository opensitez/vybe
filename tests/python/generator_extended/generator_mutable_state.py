# vybe-test: python/generator_extended/generator_mutable_state
# origin: languages/python/tests/python/test_generator_extended.rs

def g():
 acc = []
 for i in range(3):
  acc.append(i)
  yield sum(acc)
print(list(g()))
