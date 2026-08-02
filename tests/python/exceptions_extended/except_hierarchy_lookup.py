# vybe-test: python/exceptions_extended/except_hierarchy_lookup
# origin: languages/python/tests/python/test_exceptions_extended.rs

try:
 raise LookupError()
except KeyError:
 print('ke')
except LookupError:
 print('le')
