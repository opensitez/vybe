# vybe-test: python/control_flow/for_break_runtime
# origin: languages/python/tests/python/test_control_flow.rs

for i in range(10):
    if i == 3:
        break
    print(i)
