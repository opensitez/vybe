# vybe-test: python/control_flow/try_except_finally
# origin: languages/python/tests/python/test_control_flow.rs
# vybe-test-mode: compile

try:
    x = 1
except:
    pass
finally:
    print('done')
