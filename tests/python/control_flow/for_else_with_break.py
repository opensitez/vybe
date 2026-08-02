# vybe-test: python/control_flow/for_else_with_break
# origin: languages/python/tests/python/test_control_flow.rs
# vybe-test-mode: compile

for x in [1, 2, 3]:
    if x == 2:
        break
else:
    print('unreachable')
