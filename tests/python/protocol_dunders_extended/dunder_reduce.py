# vybe-test: python/protocol_dunders_extended/dunder_reduce
# origin: languages/python/tests/python/test_protocol_dunders_extended.rs
# vybe-test-mode: compile

class V:
 def __reduce__(self):
  return (V, ())
