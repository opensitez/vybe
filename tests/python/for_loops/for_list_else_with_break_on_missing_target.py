# vybe-test: python/for_loops/for_list_else_with_break_on_missing_target
# origin: languages/python/tests/python/test_for_loops.rs

for x in [4, 5, 6]:
    if x == 1:
        break
else:
    print('all checked')
