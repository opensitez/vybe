# vybe-test: python/types_introspection_extended/hash_unhashable
# origin: languages/python/tests/python/test_types_introspection_extended.rs
# vybe-test-mode: compile

try:
 hash([])
except TypeError:
 pass
