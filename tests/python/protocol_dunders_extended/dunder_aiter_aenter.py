# vybe-test: python/protocol_dunders_extended/dunder_aiter_aenter
# origin: languages/python/tests/python/test_protocol_dunders_extended.rs

class V:
 async def __aenter__(self): return self
 async def __aexit__(self, *a): pass
