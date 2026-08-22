# vybe-test: python/control_flow/while_with_continue
# origin: languages/python/tests/python/test_control_flow.rs

i = 0
while i < 10:
    i += 1
    if i == 5:
        continue
    print(i)
