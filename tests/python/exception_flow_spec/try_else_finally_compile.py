# vybe-test: python/exception_flow_spec/try_else_finally_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
def risky(*_a, **_k):
    return None
def cleanup(*_a, **_k):
    return None
def finish(*_a, **_k):
    return None

try:
    risky()
except ValueError:
    handle()
else:
    cleanup()
finally:
    finish()
