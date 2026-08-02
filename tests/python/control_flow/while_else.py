# vybe-test: python/control_flow/while_else
# origin: languages/python/tests/python/test_control_flow.rs
# vybe-test-mode: compile

i = 0
while i < 5:
    i += 1
else:
    print('done')
