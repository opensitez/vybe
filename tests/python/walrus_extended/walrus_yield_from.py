# vybe-test: python/walrus_extended/walrus_yield_from
# origin: languages/python/tests/python/test_walrus_extended.rs
# vybe-test-mode: compile

def g():
 if (x := 1):
  yield x
