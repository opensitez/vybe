# vybe-test: python/if_else/nested_if_outer_false_skips_inner
# origin: languages/python/tests/python/test_if_else.rs

a = 1
b = 3
if a > 5:
    if b < 5:
        print('inner')
print('end')
