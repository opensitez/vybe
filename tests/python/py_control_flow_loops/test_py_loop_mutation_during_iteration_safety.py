# vybe-test: python/py_control_flow_loops/test_py_loop_mutation_during_iteration_safety
# origin: languages/python/tests/python/test_py_control_flow_loops.rs

lst = [1, 2, 3, 4, 5]
# Iterating over a copy to safely mutate original
for item in lst[:]:
    if item % 2 == 0:
        lst.remove(item)

print(lst)
