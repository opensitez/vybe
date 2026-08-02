# vybe-test: python/for_loops/for_else_skipped_when_break_occurs
# origin: languages/python/tests/python/test_for_loops.rs

for x in [1, 2, 3]:
    if x == 2:
        break
else:
    print('skipped')
print('done')
