# vybe-test: python/for_loops/for_else_runs_when_loop_completes_without_break
# origin: languages/python/tests/python/test_for_loops.rs

for x in [1, 2, 3]:
    if x == 5:
        break
else:
    print('finished')
