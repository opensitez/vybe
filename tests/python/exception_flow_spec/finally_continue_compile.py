# vybe-test: python/exception_flow_spec/finally_continue_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs

for i in range(3):
    try:
        pass
    finally:
        continue
