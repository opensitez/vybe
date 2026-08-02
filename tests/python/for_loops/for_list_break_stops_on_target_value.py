# vybe-test: python/for_loops/for_list_break_stops_on_target_value
# origin: languages/python/tests/python/test_for_loops.rs

for x in [1, 2, 3, 4, 5]:
    if x == 3:
        break
    print(x)
