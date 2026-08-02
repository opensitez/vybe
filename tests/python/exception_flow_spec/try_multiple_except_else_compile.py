# vybe-test: python/exception_flow_spec/try_multiple_except_else_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    risky()
except ValueError:
    pass
except TypeError:
    pass
else:
    done()
