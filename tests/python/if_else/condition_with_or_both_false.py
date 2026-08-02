# vybe-test: python/if_else/condition_with_or_both_false
# origin: languages/python/tests/python/test_if_else.rs

a = -1
b = -2
if a > 0 or b > 0:
    print('either')
else:
    print('neither')
