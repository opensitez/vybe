# vybe-test: python/callable_hash_id/hash_dict_not_hashable_raises
# origin: languages/python/tests/python/test_callable_hash_id.rs

try:
 hash({})
 print('ok')
except TypeError:
 print('TypeError')
