# vybe-test: python/context_managers_extended/with_async_context
# origin: languages/python/tests/python/test_context_managers_extended.rs

class CM:
 async def __aenter__(self):
  return self
 async def __aexit__(self, *a):
  pass
