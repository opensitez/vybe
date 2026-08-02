# vybe-test: python/callable_hash_id/id_small_ints_may_be_cached
# origin: languages/python/tests/python/test_callable_hash_id.rs

id(256) == id(256) or id(256) != id(256)
