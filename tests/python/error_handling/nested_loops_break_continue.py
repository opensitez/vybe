# vybe-test: python/error_handling/nested_loops_break_continue
# origin: languages/python/tests/python/test_error_handling.rs

for i in range(3):
    for j in range(3):
        if j == 1:
            continue
        if i == 2:
            break
        print(i, j)
