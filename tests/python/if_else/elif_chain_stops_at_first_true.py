# vybe-test: python/if_else/elif_chain_stops_at_first_true
# origin: languages/python/tests/python/test_if_else.rs

n = 95
if n >= 90:
    print('A')
elif n >= 80:
    print('B')
elif n >= 70:
    print('C')
else:
    print('F')
