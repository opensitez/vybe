# vybe-test: python/for_loops/for_list_continue_skips_all_multiples_of_three
# origin: languages/python/tests/python/test_for_loops.rs

for n in [1, 2, 3, 4, 5, 6, 7, 8, 9]:
    if n % 3 == 0:
        continue
    print(n)
