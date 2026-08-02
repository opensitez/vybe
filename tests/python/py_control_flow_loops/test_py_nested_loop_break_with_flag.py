# vybe-test: python/py_control_flow_loops/test_py_nested_loop_break_with_flag
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

matrix = [[1, 2], [3, 4], [5, 6]]
found = None
for row in matrix:
    for val in row:
        if val == 4:
            found = val
            break
    if found is not None:
        break

print(found)
