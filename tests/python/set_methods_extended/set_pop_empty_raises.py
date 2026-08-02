# vybe-test: python/set_methods_extended/set_pop_empty_raises
# origin: languages/python/tests/python/test_set_methods_extended.rs

s = set()
try:
    s.pop()
    print('ok')
except KeyError:
    print('empty')
