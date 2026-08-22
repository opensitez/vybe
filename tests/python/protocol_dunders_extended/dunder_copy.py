# vybe-test: python/protocol_dunders_extended/dunder_copy
# origin: languages/python/tests/python/test_protocol_dunders_extended.rs

class V:
 def __copy__(self):
  return V()
