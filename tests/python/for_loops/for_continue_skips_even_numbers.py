# vybe-test: python/for_loops/for_continue_skips_even_numbers
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(6):
    if i % 2 == 0:
        continue
    print(i)
