# vybe-test: python/for_loops/for_list_continue_skips_matching_value
# origin: languages/python/tests/python/test_for_loops.rs

for x in [1, 2, 3, 4]:
    if x == 2:
        continue
    print(x)
