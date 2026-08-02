# vybe-test: python/exception_flow_spec/nested_finally_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    try:
        risky()
    finally:
        clean_inner()
finally:
    clean_outer()
