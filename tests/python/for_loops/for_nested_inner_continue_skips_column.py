# vybe-test: python/for_loops/for_nested_inner_continue_skips_column
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(3):
    for j in range(3):
        if j == 1:
            continue
        if i == 2:
            break
        print(i, j)
