# vybe-test: python/walrus_extended/walrus_with_statement
# origin: languages/python/tests/python/test_walrus_extended.rs
# vybe-test-mode: compile

class CM:
 def __enter__(self): return self
 def __exit__(self, *a): pass
with CM() as (c := CM()):
 pass
