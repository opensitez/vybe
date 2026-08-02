# vybe-test: python/walrus_core/walrus_scope_in_comprehension_local
# origin: languages/python/tests/python/test_walrus_core.rs

out = [z for _ in [1] if (z := 9)]
print(z if 'z' in dir() else out)
