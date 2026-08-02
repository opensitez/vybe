# vybe-test: python/exception_flow_spec/try_else_finally_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
# vybe-test-mode: compile

try:
    risky()
except ValueError:
    handle()
else:
    cleanup()
finally:
    finish()
