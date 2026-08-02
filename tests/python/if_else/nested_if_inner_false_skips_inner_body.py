# vybe-test: python/if_else/nested_if_inner_false_skips_inner_body
# origin: languages/python/tests/python/test_if_else.rs

a = 10
b = 9
if a > 5:
    if b < 5:
        print('inner')
    else:
        print('else-inner')
