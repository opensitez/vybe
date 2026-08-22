# vybe-test: python/protocol_dunders_extended/dunder_getnewargs
# origin: languages/python/tests/python/test_protocol_dunders_extended.rs

class V:
 def __getnewargs__(self):
  return ()
