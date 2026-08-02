# vybe-test: python/py_control_flow_loops/test_py_loop_unpacking_structures
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

pairs = [("a", 1), ("b", 2), ("c", 3)]
out = []
for k, v in pairs:
    out.append(f"{k}:{v}")

print(", ".join(out))
