# vybe-test: python/py_control_flow_loops/test_py_while_else_loop
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

count = 3
while count > 0:
    count -= 1
else:
    print("while-else reached")

count = 3
while count > 0:
    if count == 2:
        break
    count -= 1
else:
    print("while-else skipped on break")
