# vybe-test: python/if_else/condition_and_or_precedence_in_if
# origin: languages/python/tests/python/test_if_else.rs

a = True
b = False
c = True
if a or b and c:
    print('yes')
else:
    print('no')
