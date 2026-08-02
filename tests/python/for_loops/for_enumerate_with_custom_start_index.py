# vybe-test: python/for_loops/for_enumerate_with_custom_start_index
# origin: languages/python/tests/python/test_for_loops.rs

for i, item in enumerate(['x', 'y'], start=10):
    print(i, item)
