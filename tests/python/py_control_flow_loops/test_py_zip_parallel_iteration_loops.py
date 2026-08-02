# vybe-test: python/py_control_flow_loops/test_py_zip_parallel_iteration_loops
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

names = ["Alice", "Bob", "Charlie"]
ages = [25, 30, 35]

for name, age in zip(names, ages):
    print(f"{name} is {age}")
