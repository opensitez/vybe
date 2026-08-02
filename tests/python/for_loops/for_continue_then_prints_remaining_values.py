# vybe-test: python/for_loops/for_continue_then_prints_remaining_values
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(5):
    if i == 2:
        continue
    print(i)
