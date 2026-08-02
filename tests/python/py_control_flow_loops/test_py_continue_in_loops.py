# vybe-test: python/py_control_flow_loops/test_py_continue_in_loops
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

evens = []
for i in range(10):
    if i % 2 != 0:
        continue
    evens.append(i)

print(evens)
