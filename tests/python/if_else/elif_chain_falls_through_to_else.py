# vybe-test: python/if_else/elif_chain_falls_through_to_else
# origin: languages/python/tests/python/test_if_else.rs

score = 55
if score >= 90:
    print('A')
elif score >= 70:
    print('B')
else:
    print('C')
