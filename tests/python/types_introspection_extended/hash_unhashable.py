# vybe-test: python/types_introspection_extended/hash_unhashable
# origin: languages/python/tests/python/test_types_introspection_extended.rs

try:
 hash([])
except TypeError:
 pass
