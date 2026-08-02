# vybe-test: python/oop_inheritance_extended/dunder_hash_none_by_default
# origin: languages/python/tests/python/test_oop_inheritance_extended.rs

class C:
 pass
try:
 hash(C())
 print('h')
except TypeError:
 print('unhashable')
