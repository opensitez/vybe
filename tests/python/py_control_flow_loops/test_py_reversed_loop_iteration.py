# vybe-test: python/py_control_flow_loops/test_py_reversed_loop_iteration
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

items = ["first", "second", "third"]
out = []
for item in reversed(items):
    out.append(item)

print(out)
