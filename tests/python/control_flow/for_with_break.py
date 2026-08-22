# vybe-test: python/control_flow/for_with_break
# origin: languages/python/tests/python/test_control_flow.rs

for i in range(10):
    if i == 5:
        break
    print(i)
