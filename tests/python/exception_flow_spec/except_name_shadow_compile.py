# vybe-test: python/exception_flow_spec/except_name_shadow_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    risky()
except ValueError as value_error:
    print(value_error)
