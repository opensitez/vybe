# vybe-test: python/control_flow/for_continue_runtime
# origin: languages/python/tests/python/test_control_flow.rs

for i in range(5):
    if i == 2:
        continue
    print(i)
