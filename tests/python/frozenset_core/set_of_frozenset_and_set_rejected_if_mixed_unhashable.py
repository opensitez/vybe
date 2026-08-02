# vybe-test: python/frozenset_core/set_of_frozenset_and_set_rejected_if_mixed_unhashable
# origin: languages/python/tests/python/test_frozenset_core.rs

try:
 {frozenset([1]), [2]}
 print('ok')
except TypeError:
 print('TypeError')
