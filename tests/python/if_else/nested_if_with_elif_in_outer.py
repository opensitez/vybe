# vybe-test: python/if_else/nested_if_with_elif_in_outer
# origin: languages/python/tests/python/test_if_else.rs

x = 5
y = 2
if x > 10:
    print('outer-a')
elif x > 3:
    if y < 5:
        print('inner')
else:
    print('outer-b')
