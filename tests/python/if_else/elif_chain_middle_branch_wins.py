# vybe-test: python/if_else/elif_chain_middle_branch_wins
# origin: languages/python/tests/python/test_if_else.rs

n = 85
if n >= 90:
    print('A')
elif n >= 80:
    print('B')
elif n >= 70:
    print('C')
else:
    print('F')
