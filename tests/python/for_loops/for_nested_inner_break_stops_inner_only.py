# vybe-test: python/for_loops/for_nested_inner_break_stops_inner_only
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(2):
    for j in range(3):
        if j == 1:
            break
        print(i, j)
