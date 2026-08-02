# vybe-test: python/for_loops/for_enumerate_break_midway
# origin: languages/python/tests/python/test_for_loops.rs

for i, val in enumerate([10, 20, 30, 40]):
    if i == 2:
        break
    print(val)
