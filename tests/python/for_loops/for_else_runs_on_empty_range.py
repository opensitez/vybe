# vybe-test: python/for_loops/for_else_runs_on_empty_range
# origin: languages/python/tests/python/test_for_loops.rs

for x in range(0):
    print(x)
else:
    print('empty')
