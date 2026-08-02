# vybe-test: python/for_loops/for_range_break_before_any_output
# origin: languages/python/tests/python/test_for_loops.rs

for i in range(5):
    if i == 0:
        break
    print(i)
print('end')
