# vybe-test: python/if_else/elif_chain_picks_first_matching_branch
# origin: languages/python/tests/python/test_if_else.rs

score = 75
if score >= 90:
    print('A')
elif score >= 70:
    print('B')
else:
    print('C')
