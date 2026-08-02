# vybe-test: python/if_else/nested_if_else_if_chain
# origin: languages/python/tests/python/test_if_else.rs

x = 15
if x < 10:
    print('low')
else:
    if x < 20:
        print('mid')
    else:
        print('high')
