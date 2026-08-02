# vybe-test: python/control_flow/nested_loops
# origin: languages/python/tests/python/test_control_flow.rs
# vybe-test-mode: compile

for i in range(3):
    for j in range(3):
        print(i, j)
