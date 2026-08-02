# vybe-test: python/protocol_dunders_extended/dunder_await
# origin: languages/python/tests/python/test_protocol_dunders_extended.rs
# vybe-test-mode: compile

class V:
 def __await__(self):
  yield 1
