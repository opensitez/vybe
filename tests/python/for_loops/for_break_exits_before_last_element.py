# vybe-test: python/for_loops/for_break_exits_before_last_element
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(10):
    if i == 3:
        break
    print(i)
