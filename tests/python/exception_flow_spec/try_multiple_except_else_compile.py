# vybe-test: python/exception_flow_spec/try_multiple_except_else_compile
# origin: languages/python/tests/python/test_exception_flow_spec.rs
def risky(*_a, **_k):
    return None
def done(*_a, **_k):
    return None

try:
    risky()
except ValueError:
    pass
except TypeError:
    pass
else:
    done()
